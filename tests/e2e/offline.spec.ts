import { cp, readFile, writeFile } from 'node:fs/promises';
import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { extname, join } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { excelWorkbook } from '../xlsx';
import { clickCell, enter, expectCell, openApp, openFile } from './grid';

test.use({ locale: 'en-US' });

/** Waits until the service worker has taken the page over. */
async function workingOffline(page: Page): Promise<void> {
  await page.waitForFunction(
    async () =>
      (await navigator.serviceWorker.ready) &&
      navigator.serviceWorker.controller !== null,
  );
}

test('after one visit, the app, its stored workbooks and its files work offline', async ({
  page,
  context,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '42');
  await workingOffline(page);

  await context.setOffline(true);
  await page.goto('./');
  await page.getByRole('button', { name: /^Open Book1/ }).click();
  await expectCell(page, 0, 0, 'A1 42');

  await openFile(page, 'Excel.xlsx', await excelWorkbook());
  await expect(page.getByLabel('Workbook name')).toHaveValue('Excel');
  await page.getByRole('tab', { name: 'File' }).click();
  const saving = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save' }).click();
  expect((await saving).suggestedFilename()).toBe('Excel.xlsx');

  await openFile(
    page,
    'Prices.csv',
    new TextEncoder().encode('Item,Price\r\nTea,3.75\r\n'),
    'text/csv',
  );
  await expectCell(page, 1, 1, 'B2 3.75');
  await page.getByRole('tab', { name: 'File' }).click();
  const exporting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Export CSV' }).click();
  const exported = await readFile(await (await exporting).path());
  expect(exported.toString('utf8')).toBe('﻿Item,Price\r\nTea,3.75\r\n');
});

const TYPES: Record<string, string> = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
};

/** Serves a folder at `/gridworks/`, as GitHub Pages serves the build. */
async function serve(folder: string): Promise<Server> {
  const server = createServer((request, response) => {
    const path = new URL(request.url!, 'http://localhost').pathname;
    const file = path.replace(/^\/gridworks\//, '') || 'index.html';
    readFile(join(folder, file)).then(
      (bytes) => {
        response.writeHead(200, {
          'Content-Type': TYPES[extname(file)] ?? 'application/octet-stream',
          'Cache-Control': 'no-cache',
        });
        response.end(bytes);
      },
      () => {
        response.writeHead(404);
        response.end();
      },
    );
  });
  await new Promise<void>((resolve) => server.listen(0, resolve));
  return server;
}

test('a new version deployed is loaded by the next visit, which asks for a reload', async ({
  page,
}, testInfo) => {
  // A copy of the build, which this test deploys again with a change.
  const site = testInfo.outputPath('site');
  await cp('dist', site, { recursive: true });
  const server = await serve(site);
  try {
    const { port } = server.address() as AddressInfo;
    const home = `http://localhost:${port}/gridworks/`;
    await page.goto(home);
    await workingOffline(page);
    const reload = page.getByRole('alert');
    await expect(reload).toBeHidden();

    // The new version: another page, whose revision the worker lists.
    const index = join(site, 'index.html');
    await writeFile(
      index,
      (await readFile(index, 'utf8')).replace(
        '</head>',
        '<meta name="version" content="2"></head>',
      ),
    );
    const worker = join(site, 'sw.js');
    await writeFile(
      worker,
      (await readFile(worker, 'utf8')).replace(
        /("?url"?:\s*"index\.html",\s*"?revision"?:\s*")[^"]+/,
        '$1second',
      ),
    );

    await page.goto(home);
    await expect(reload).toHaveText(
      'A new version of Gridworks is ready. Reload the page to use it.',
    );
    await page.reload();
    await expect(page.locator('meta[name="version"]')).toHaveCount(1);
    await expect(reload).toBeHidden();
  } finally {
    server.close();
  }
});
