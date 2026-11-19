/**
 * The descriptions of the functions of the library and of the values
 * Excel offers for some of their arguments, as Excel's list shows them,
 * by the names of the engine's functions, arguments and values.
 */
export const descriptionsEn = {
  'function.ABS':
    'Returns the absolute value of a number, a number without its sign',
  'function.ACOS':
    'Returns the arccosine of a number, in radians in the range 0 to Pi',
  'function.ACOSH': 'Returns the inverse hyperbolic cosine of a number',
  'function.ADDRESS':
    'Creates a cell reference as text, given specified row and column numbers',
  'function.AND':
    'Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE',
  'function.ASIN':
    'Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2',
  'function.ASINH': 'Returns the inverse hyperbolic sine of a number',
  'function.ATAN':
    'Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2',
  'function.ATAN2':
    'Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi',
  'function.ATANH': 'Returns the inverse hyperbolic tangent of a number',
  'function.AVEDEV':
    'Returns the average of the absolute deviations of data points from their mean. Arguments can be numbers or names, arrays, or references that contain numbers',
  'function.AVERAGE':
    'Returns the average (arithmetic mean) of its arguments, which can be numbers or names, arrays, or references that contain numbers',
  'function.AVERAGEA':
    'Returns the average (arithmetic mean) of its arguments, evaluating text and FALSE in arguments as 0; TRUE evaluates as 1',
  'function.AVERAGEIF':
    'Finds average (arithmetic mean) for the cells specified by a given condition or criteria',
  'function.AVERAGEIFS':
    'Finds average (arithmetic mean) for the cells specified by a given set of conditions or criteria',
  'function.CEILING':
    'Rounds a number up, to the nearest multiple of significance',
  'function.CEILING.MATH':
    'Rounds a number up, to the nearest integer or to the nearest multiple of significance',
  'function.CHAR':
    'Returns the character specified by the code number from the character set for your computer',
  'function.CHOOSE':
    'Chooses a value or action to perform from a list of values, based on an index number',
  'function.CLEAN': 'Removes all nonprintable characters from text',
  'function.CODE':
    'Returns a numeric code for the first character in a text string, in the character set used by your computer',
  'function.COLUMN': 'Returns the column number of a reference',
  'function.COLUMNS': 'Returns the number of columns in an array or reference',
  'function.COMBIN':
    'Returns the number of combinations for a given number of items',
  'function.CONCAT': 'Concatenates a list or range of text strings',
  'function.CONCATENATE': 'Joins several text strings into one text string',
  'function.CORREL':
    'Returns the correlation coefficient between two data sets',
  'function.COS': 'Returns the cosine of an angle',
  'function.COSH': 'Returns the hyperbolic cosine of a number',
  'function.COUNT':
    'Counts the number of cells in a range that contain numbers',
  'function.COUNTA': 'Counts the number of cells in a range that are not empty',
  'function.COUNTBLANK':
    'Counts the number of empty cells in a specified range of cells',
  'function.COUNTIF':
    'Counts the number of cells within a range that meet the given condition',
  'function.COUNTIFS':
    'Counts the number of cells specified by a given set of conditions or criteria',
  'function.DATE':
    'Returns the number that represents the date in Microsoft Excel date-time code',
  'function.DATEDIF':
    'Calculates the number of days, months, or years between two dates',
  'function.DATEVALUE':
    'Converts a date in the form of text to a number that represents the date in Microsoft Excel date-time code',
  'function.DAY': 'Returns the day of the month, a number from 1 to 31',
  'function.DAYS': 'Returns the number of days between the two dates',
  'function.DAYS360':
    'Returns the number of days between two dates based on a 360-day year (twelve 30-day months)',
  'function.DEGREES': 'Converts radians to degrees',
  'function.DEVSQ':
    'Returns the sum of squares of deviations of data points from their sample mean',
  'function.DOLLAR': 'Converts a number to text, using currency format',
  'function.EDATE':
    'Returns the serial number of the date that is the indicated number of months before or after the start date',
  'function.EOMONTH':
    'Returns the serial number of the last day of the month before or after a specified number of months',
  'function.ERROR.TYPE': 'Returns a number matching an error value',
  'function.EVEN':
    'Rounds a positive number up and negative number down to the nearest even integer',
  'function.EXACT':
    'Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive',
  'function.EXP': 'Returns e raised to the power of a given number',
  'function.FACT':
    'Returns the factorial of a number, equal to 1*2*3*...* Number',
  'function.FACTDOUBLE': 'Returns the double factorial of a number',
  'function.FALSE': 'Returns the logical value FALSE',
  'function.FIND':
    'Returns the starting position of one text string within another text string. FIND is case-sensitive',
  'function.FIXED':
    'Rounds a number to the specified number of decimals and returns the result as text with or without commas',
  'function.FLOOR':
    'Rounds a number down to the nearest multiple of significance',
  'function.FLOOR.MATH':
    'Rounds a number down, to the nearest integer or to the nearest multiple of significance',
  'function.FORMULATEXT': 'Returns a formula as a string',
  'function.GCD': 'Returns the greatest common divisor',
  'function.GEOMEAN':
    'Returns the geometric mean of an array or range of positive numeric data',
  'function.HARMEAN':
    'Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals',
  'function.HLOOKUP':
    'Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify',
  'function.HOUR':
    'Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.)',
  'function.IF':
    'Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE',
  'function.IFERROR':
    'Returns value_if_error if expression is an error and the value of the expression itself otherwise',
  'function.IFNA':
    'Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression',
  'function.IFS':
    'Checks whether one or more conditions are met and returns a value corresponding to the first TRUE condition',
  'function.INDEX':
    'Returns a value or reference of the cell at the intersection of a particular row and column, in a given range',
  'function.INT': 'Rounds a number down to the nearest integer',
  'function.ISBLANK':
    'Checks whether a reference is to an empty cell, and returns TRUE or FALSE',
  'function.ISERR':
    'Checks whether a value is an error other than #N/A, and returns TRUE or FALSE',
  'function.ISERROR':
    'Checks whether a value is an error, and returns TRUE or FALSE',
  'function.ISEVEN': 'Returns TRUE if the number is even',
  'function.ISFORMULA':
    'Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE',
  'function.ISLOGICAL':
    'Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE',
  'function.ISNA': 'Checks whether a value is #N/A, and returns TRUE or FALSE',
  'function.ISNONTEXT':
    'Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE',
  'function.ISNUMBER':
    'Checks whether a value is a number, and returns TRUE or FALSE',
  'function.ISODD': 'Returns TRUE if the number is odd',
  'function.ISOWEEKNUM':
    'Returns the ISO week number in the year for a given date',
  'function.ISREF':
    'Checks whether a value is a reference, and returns TRUE or FALSE',
  'function.ISTEXT':
    'Checks whether a value is text, and returns TRUE or FALSE',
  'function.LARGE':
    'Returns the k-th largest value in a data set. For example, the fifth largest number',
  'function.LCM': 'Returns the least common multiple',
  'function.LEFT':
    'Returns the specified number of characters from the start of a text string',
  'function.LEN': 'Returns the number of characters in a text string',
  'function.LN': 'Returns the natural logarithm of a number',
  'function.LOG': 'Returns the logarithm of a number to the base you specify',
  'function.LOG10': 'Returns the base-10 logarithm of a number',
  'function.LOOKUP':
    'Looks up a value either from a one-row or one-column range or from an array. Provided for backward compatibility',
  'function.LOWER': 'Converts all letters in a text string to lowercase',
  'function.MATCH':
    'Returns the relative position of an item in an array that matches a specified value in a specified order',
  'function.MAX':
    'Returns the largest value in a set of values. Ignores logical values and text',
  'function.MAXA':
    'Returns the largest value in a set of values. Does not ignore logical values and text',
  'function.MAXIFS':
    'Returns the maximum value among cells specified by a given set of conditions or criteria',
  'function.MEDIAN':
    'Returns the median, or the number in the middle of the set of given numbers',
  'function.MID':
    'Returns the characters from the middle of a text string, given a starting position and length',
  'function.MIN':
    'Returns the smallest number in a set of values. Ignores logical values and text',
  'function.MINA':
    'Returns the smallest value in a set of values. Does not ignore logical values and text',
  'function.MINIFS':
    'Returns the minimum value among cells specified by a given set of conditions or criteria',
  'function.MINUTE': 'Returns the minute, a number from 0 to 59',
  'function.MOD':
    'Returns the remainder after a number is divided by a divisor',
  'function.MODE':
    'Returns the most frequently occurring, or repetitive, value in an array or range of data',
  'function.MODE.SNGL':
    'Returns the most frequently occurring, or repetitive, value in an array or range of data',
  'function.MONTH':
    'Returns the month, a number from 1 (January) to 12 (December)',
  'function.MROUND': 'Returns a number rounded to the desired multiple',
  'function.N':
    'Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero)',
  'function.NA': 'Returns the error value #N/A (value not available)',
  'function.NETWORKDAYS':
    'Returns the number of whole workdays between two dates',
  'function.NETWORKDAYS.INTL':
    'Returns the number of whole workdays between two dates with custom weekend parameters',
  'function.NOT': 'Changes FALSE to TRUE, or TRUE to FALSE',
  'function.NOW':
    'Returns the current date and time formatted as a date and time',
  'function.NUMBERVALUE':
    'Converts text to number in a locale-independent manner',
  'function.ODD':
    'Rounds a positive number up and negative number down to the nearest odd integer',
  'function.OR':
    'Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE',
  'function.PERCENTILE': 'Returns the k-th percentile of values in a range',
  'function.PERCENTILE.INC':
    'Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive',
  'function.PI':
    'Returns the value of Pi, 3.14159265358979, accurate to 15 digits',
  'function.POWER': 'Returns the result of a number raised to a power',
  'function.PRODUCT': 'Multiplies all the numbers given as arguments',
  'function.PROPER':
    'Converts a text string to proper case; the first letter in each word to uppercase, and all other letters to lowercase',
  'function.QUARTILE': 'Returns the quartile of a data set',
  'function.QUARTILE.INC':
    'Returns the quartile of a data set, based on percentile values from 0..1, inclusive',
  'function.QUOTIENT': 'Returns the integer portion of a division',
  'function.RADIANS': 'Converts degrees to radians',
  'function.RANK':
    'Returns the rank of a number in a list of numbers: its size relative to other values in the list',
  'function.RANK.AVG':
    'Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the average rank is returned',
  'function.RANK.EQ':
    'Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the top rank of that set of values is returned',
  'function.REPLACE':
    'Replaces part of a text string with a different text string',
  'function.REPT':
    'Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string',
  'function.RIGHT':
    'Returns the specified number of characters from the end of a text string',
  'function.ROUND': 'Rounds a number to a specified number of digits',
  'function.ROUNDDOWN': 'Rounds a number down, toward zero',
  'function.ROUNDUP': 'Rounds a number up, away from zero',
  'function.ROW': 'Returns the row number of a reference',
  'function.ROWS': 'Returns the number of rows in a reference or array',
  'function.SEARCH':
    'Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive)',
  'function.SECOND': 'Returns the second, a number from 0 to 59',
  'function.SIGN':
    'Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative',
  'function.SIN': 'Returns the sine of an angle',
  'function.SINH': 'Returns the hyperbolic sine of a number',
  'function.SMALL':
    'Returns the k-th smallest value in a data set. For example, the fifth smallest number',
  'function.SQRT': 'Returns the square root of a number',
  'function.SQRTPI': 'Returns the square root of (number * Pi)',
  'function.STDEV':
    'Estimates standard deviation based on a sample (ignores logical values and text in the sample)',
  'function.STDEV.P':
    'Calculates standard deviation based on the entire population given as arguments (ignores logical values and text)',
  'function.STDEV.S':
    'Estimates standard deviation based on a sample (ignores logical values and text in the sample)',
  'function.STDEVP':
    'Calculates standard deviation based on the entire population given as arguments (ignores logical values and text)',
  'function.SUBSTITUTE':
    'Replaces existing text with new text in a text string',
  'function.SUM': 'Adds all the numbers in a range of cells',
  'function.SUMIF': 'Adds the cells specified by a given condition or criteria',
  'function.SUMIFS':
    'Adds the cells specified by a given set of conditions or criteria',
  'function.SUMPRODUCT':
    'Returns the sum of the products of corresponding ranges or arrays',
  'function.SUMSQ':
    'Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names, or references to cells that contain numbers',
  'function.SWITCH':
    'Evaluates an expression against a list of values and returns the result corresponding to the first matching value. If there is no match, an optional default value is returned',
  'function.T':
    'Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not',
  'function.TAN': 'Returns the tangent of an angle',
  'function.TANH': 'Returns the hyperbolic tangent of a number',
  'function.TEXT': 'Converts a value to text in a specific number format',
  'function.TEXTJOIN':
    'Concatenates a list or range of text strings using a delimiter',
  'function.TIME':
    'Converts hours, minutes, and seconds given as numbers to an Excel serial number, formatted with a time format',
  'function.TIMEVALUE':
    'Converts a text time to an Excel serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM)',
  'function.TODAY': 'Returns the current date formatted as a date',
  'function.TRIM':
    'Removes all spaces from a text string except for single spaces between words',
  'function.TRUE': 'Returns the logical value TRUE',
  'function.TRUNC':
    'Truncates a number to an integer by removing the decimal, or fractional, part of the number',
  'function.TYPE':
    'Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64',
  'function.UNICHAR':
    'Returns the Unicode character referenced by the given numeric value',
  'function.UNICODE':
    'Returns the number (code point) corresponding to the first character of the text',
  'function.UPPER': 'Converts a text string to all uppercase letters',
  'function.VALUE':
    'Converts a text string that represents a number to a number',
  'function.VAR':
    'Estimates variance based on a sample (ignores logical values and text in the sample)',
  'function.VAR.P':
    'Calculates variance based on the entire population (ignores logical values and text in the population)',
  'function.VAR.S':
    'Estimates variance based on a sample (ignores logical values and text in the sample)',
  'function.VARP':
    'Calculates variance based on the entire population (ignores logical values and text in the population)',
  'function.VLOOKUP':
    'Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify. By default, the table must be sorted in an ascending order',
  'function.WEEKDAY':
    'Returns a number from 1 to 7 identifying the day of the week of a date',
  'function.WEEKNUM': 'Returns the week number in the year',
  'function.WORKDAY':
    'Returns the serial number of the date before or after a specified number of workdays',
  'function.WORKDAY.INTL':
    'Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters',
  'function.XLOOKUP':
    'Searches a range or an array for a match and returns the corresponding item from a second range or array. By default, an exact match is used',
  'function.XMATCH':
    'Returns the relative position of an item in an array. By default, an exact match is required',
  'function.XOR': "Returns a logical 'Exclusive Or' of all arguments",
  'function.YEAR':
    'Returns the year of a date, an integer in the range 1900 - 9999',
  'function.YEARFRAC':
    'Returns the year fraction representing the number of whole days between start_date and end_date',
  'value.ADDRESS.abs_num.1': 'Absolute',
  'value.ADDRESS.abs_num.2': 'Absolute row; relative column',
  'value.ADDRESS.abs_num.3': 'Relative row; absolute column',
  'value.ADDRESS.abs_num.4': 'Relative',
  'value.ADDRESS.a1.TRUE': 'A1 style',
  'value.ADDRESS.a1.FALSE': 'R1C1 style',
  'value.DAYS360.method.FALSE': 'U.S. (NASD) method',
  'value.DAYS360.method.TRUE': 'European method',
  'value.HLOOKUP.range_lookup.TRUE': 'Approximate match',
  'value.HLOOKUP.range_lookup.FALSE': 'Exact match',
  'value.MATCH.match_type.1': 'Less than',
  'value.MATCH.match_type.0': 'Exact match',
  'value.MATCH.match_type.-1': 'Greater than',
  'value.NETWORKDAYS.INTL.weekend.1': 'Saturday, Sunday',
  'value.NETWORKDAYS.INTL.weekend.2': 'Sunday, Monday',
  'value.NETWORKDAYS.INTL.weekend.3': 'Monday, Tuesday',
  'value.NETWORKDAYS.INTL.weekend.4': 'Tuesday, Wednesday',
  'value.NETWORKDAYS.INTL.weekend.5': 'Wednesday, Thursday',
  'value.NETWORKDAYS.INTL.weekend.6': 'Thursday, Friday',
  'value.NETWORKDAYS.INTL.weekend.7': 'Friday, Saturday',
  'value.NETWORKDAYS.INTL.weekend.11': 'Sunday only',
  'value.NETWORKDAYS.INTL.weekend.12': 'Monday only',
  'value.NETWORKDAYS.INTL.weekend.13': 'Tuesday only',
  'value.NETWORKDAYS.INTL.weekend.14': 'Wednesday only',
  'value.NETWORKDAYS.INTL.weekend.15': 'Thursday only',
  'value.NETWORKDAYS.INTL.weekend.16': 'Friday only',
  'value.NETWORKDAYS.INTL.weekend.17': 'Saturday only',
  'value.RANK.order.0': 'Descending',
  'value.RANK.order.1': 'Ascending',
  'value.RANK.AVG.order.0': 'Descending',
  'value.RANK.AVG.order.1': 'Ascending',
  'value.RANK.EQ.order.0': 'Descending',
  'value.RANK.EQ.order.1': 'Ascending',
  'value.VLOOKUP.range_lookup.TRUE': 'Approximate match',
  'value.VLOOKUP.range_lookup.FALSE': 'Exact match',
  'value.WEEKDAY.return_type.1': 'Numbers 1 (Sunday) through 7 (Saturday)',
  'value.WEEKDAY.return_type.2': 'Numbers 1 (Monday) through 7 (Sunday)',
  'value.WEEKDAY.return_type.3': 'Numbers 0 (Monday) through 6 (Sunday)',
  'value.WEEKDAY.return_type.11': 'Numbers 1 (Monday) through 7 (Sunday)',
  'value.WEEKDAY.return_type.12': 'Numbers 1 (Tuesday) through 7 (Monday)',
  'value.WEEKDAY.return_type.13': 'Numbers 1 (Wednesday) through 7 (Tuesday)',
  'value.WEEKDAY.return_type.14': 'Numbers 1 (Thursday) through 7 (Wednesday)',
  'value.WEEKDAY.return_type.15': 'Numbers 1 (Friday) through 7 (Thursday)',
  'value.WEEKDAY.return_type.16': 'Numbers 1 (Saturday) through 7 (Friday)',
  'value.WEEKDAY.return_type.17': 'Numbers 1 (Sunday) through 7 (Saturday)',
  'value.WEEKNUM.return_type.1': 'Week begins on Sunday',
  'value.WEEKNUM.return_type.2': 'Week begins on Monday',
  'value.WEEKNUM.return_type.11': 'Week begins on Monday',
  'value.WEEKNUM.return_type.12': 'Week begins on Tuesday',
  'value.WEEKNUM.return_type.13': 'Week begins on Wednesday',
  'value.WEEKNUM.return_type.14': 'Week begins on Thursday',
  'value.WEEKNUM.return_type.15': 'Week begins on Friday',
  'value.WEEKNUM.return_type.16': 'Week begins on Saturday',
  'value.WEEKNUM.return_type.17': 'Week begins on Sunday',
  'value.WEEKNUM.return_type.21': 'Week begins on Monday (System 2)',
  'value.WORKDAY.INTL.weekend.1': 'Saturday, Sunday',
  'value.WORKDAY.INTL.weekend.2': 'Sunday, Monday',
  'value.WORKDAY.INTL.weekend.3': 'Monday, Tuesday',
  'value.WORKDAY.INTL.weekend.4': 'Tuesday, Wednesday',
  'value.WORKDAY.INTL.weekend.5': 'Wednesday, Thursday',
  'value.WORKDAY.INTL.weekend.6': 'Thursday, Friday',
  'value.WORKDAY.INTL.weekend.7': 'Friday, Saturday',
  'value.WORKDAY.INTL.weekend.11': 'Sunday only',
  'value.WORKDAY.INTL.weekend.12': 'Monday only',
  'value.WORKDAY.INTL.weekend.13': 'Tuesday only',
  'value.WORKDAY.INTL.weekend.14': 'Wednesday only',
  'value.WORKDAY.INTL.weekend.15': 'Thursday only',
  'value.WORKDAY.INTL.weekend.16': 'Friday only',
  'value.WORKDAY.INTL.weekend.17': 'Saturday only',
  'value.XLOOKUP.match_mode.0': 'Exact match',
  'value.XLOOKUP.match_mode.-1': 'Exact match or next smaller item',
  'value.XLOOKUP.match_mode.1': 'Exact match or next larger item',
  'value.XLOOKUP.match_mode.2': 'Wildcard character match',
  'value.XLOOKUP.search_mode.1': 'Search first-to-last',
  'value.XLOOKUP.search_mode.-1': 'Search last-to-first',
  'value.XLOOKUP.search_mode.2': 'Binary search (sorted ascending order)',
  'value.XLOOKUP.search_mode.-2': 'Binary search (sorted descending order)',
  'value.XMATCH.match_mode.0': 'Exact match',
  'value.XMATCH.match_mode.-1': 'Exact match or next smaller item',
  'value.XMATCH.match_mode.1': 'Exact match or next larger item',
  'value.XMATCH.match_mode.2': 'Wildcard character match',
  'value.XMATCH.search_mode.1': 'Search first-to-last',
  'value.XMATCH.search_mode.-1': 'Search last-to-first',
  'value.XMATCH.search_mode.2': 'Binary search (sorted ascending order)',
  'value.XMATCH.search_mode.-2': 'Binary search (sorted descending order)',
  'value.YEARFRAC.basis.0': 'US (NASD) 30/360',
  'value.YEARFRAC.basis.1': 'Actual/actual',
  'value.YEARFRAC.basis.2': 'Actual/360',
  'value.YEARFRAC.basis.3': 'Actual/365',
  'value.YEARFRAC.basis.4': 'European 30/360',
};

export const descriptionsZh: Record<keyof typeof descriptionsEn, string> = {
  'function.ABS': '返回给定数值的绝对值，即不带符号的数值',
  'function.ACOS': '返回一个弧度的反余弦。弧度值在 0 到 Pi 之间',
  'function.ACOSH': '返回反双曲余弦值',
  'function.ADDRESS': '创建一个以文本方式对工作簿中某一单元格的引用',
  'function.AND':
    '检查是否所有参数均为 TRUE，如果所有参数值均为 TRUE，则返回 TRUE',
  'function.ASIN': '返回一个弧度的反正弦。弧度值在 -Pi/2 到 Pi/2 之间',
  'function.ASINH': '返回反双曲正弦值',
  'function.ATAN': '返回反正切值。以弧度表示，大小在 -Pi/2 到 Pi/2 之间',
  'function.ATAN2':
    '根据给定的 X 轴及 Y 轴坐标值，返回反正切值。返回值在 -Pi 到 Pi 之间(不包括 -Pi)',
  'function.ATANH': '返回反双曲正切值',
  'function.AVEDEV':
    '返回一组数据点到其算术平均值的绝对偏差的平均值。参数可以是数字、名称、数组或包含数字的引用',
  'function.AVERAGE':
    '返回其参数的算术平均值；参数可以是数值或包含数值的名称、数组或引用',
  'function.AVERAGEA':
    '返回所有参数的算术平均值。字符串和 FALSE 相当于 0；TRUE 相当于 1。参数可以是数值、名称、数组或引用',
  'function.AVERAGEIF': '查找给定条件指定的单元格的平均值(算术平均值)',
  'function.AVERAGEIFS': '查找一组给定条件指定的单元格的平均值(算术平均值)',
  'function.CEILING': '将参数向上舍入为最接近的指定基数的倍数',
  'function.CEILING.MATH':
    '将数字向上舍入到最接近的整数或最接近的指定基数的倍数',
  'function.CHAR': '根据本机中的字符集，返回由代码数字指定的字符',
  'function.CHOOSE': '根据给定的索引值，从参数串中选出相应值或操作',
  'function.CLEAN': '删除文本中的所有非打印字符',
  'function.CODE': '返回文本字符串第一个字符在本机所用字符集中的数字代码',
  'function.COLUMN': '返回一引用的列号',
  'function.COLUMNS': '返回某一引用或数组的列数',
  'function.COMBIN': '返回从给定元素数目的集合中提取若干元素的组合数',
  'function.CONCAT': '连接列表或文本字符串区域',
  'function.CONCATENATE': '将多个文本字符串合并成一个',
  'function.CORREL': '返回两组数值的相关系数',
  'function.COS': '返回给定角度的余弦值',
  'function.COSH': '返回数字的双曲余弦值',
  'function.COUNT': '计算区域中包含数字的单元格的个数',
  'function.COUNTA': '计算区域中非空单元格的个数',
  'function.COUNTBLANK': '计算某个区域中空单元格的数目',
  'function.COUNTIF': '计算某个区域中满足给定条件的单元格数目',
  'function.COUNTIFS': '统计一组给定条件所指定的单元格数',
  'function.DATE': '返回在 Microsoft Excel 日期时间格式中代表日期的数字',
  'function.DATEDIF': '计算两个日期之间相隔的天数、月数或年数',
  'function.DATEVALUE':
    '将日期值从字符串转化为序列数，表示 Microsoft Excel 日期时间格式中的日期',
  'function.DAY': '返回一个月中的第几天的数值，介于 1 到 31 之间',
  'function.DAYS': '返回两个日期之间的天数',
  'function.DAYS360': '按每年 360 天返回两个日期间相差的天数(每月 30 天)',
  'function.DEGREES': '将弧度转换成角度',
  'function.DEVSQ': '返回各数据点与数据均值点之差(数据偏差)的平方和',
  'function.DOLLAR': '按照货币格式及给定的小数位数，将数字转换成文本',
  'function.EDATE': '返回一串日期，指示起始日期之前/之后的月数',
  'function.EOMONTH': '返回一串日期，表示指定月数之前或之后的月份的最后一天',
  'function.ERROR.TYPE': '返回与错误值对应的数字',
  'function.EVEN': '将正数向上舍入到最近的偶数，负数向下舍入到最近的偶数',
  'function.EXACT':
    '比较两个字符串是否完全相同(区分大小写)。返回 TRUE 或 FALSE',
  'function.EXP': '返回 e 的 n 次方',
  'function.FACT': '返回某数的阶乘，等于 1*2*3*...*Number',
  'function.FACTDOUBLE': '返回数字的双阶乘',
  'function.FALSE': '返回逻辑值 FALSE',
  'function.FIND': '返回一个字符串在另一个字符串中出现的起始位置(区分大小写)',
  'function.FIXED':
    '用定点小数格式将数值舍入成特定位数并返回带或不带逗号的文本',
  'function.FLOOR': '将参数向下舍入为最接近的指定基数的倍数',
  'function.FLOOR.MATH': '将数字向下舍入到最接近的整数或最接近的指定基数的倍数',
  'function.FORMULATEXT': '作为字符串返回公式',
  'function.GCD': '返回最大公约数',
  'function.GEOMEAN': '返回一正数数组或数值区域的几何平均数',
  'function.HARMEAN': '返回一组正数的调和平均数：所有参数倒数平均值的倒数',
  'function.HLOOKUP':
    '搜索数组区域首行满足条件的元素，确定待检索单元格在区域中的列序号，再进一步返回选定单元格的值',
  'function.HOUR':
    '返回小时数值，是一个 0 (12:00 A.M.) 到 23 (11:00 P.M.) 之间的整数',
  'function.IF':
    '判断是否满足某个条件，如果满足返回一个值，如果不满足则返回另一个值',
  'function.IFERROR':
    '如果表达式是一个错误，则返回 value_if_error，否则返回表达式自身的值',
  'function.IFNA':
    '如果表达式解析为 #N/A，则返回您指定的值，否则返回表达式的结果',
  'function.IFS': '检查是否满足一个或多个条件并返回与第一个 TRUE 条件对应的值',
  'function.INDEX': '在给定的单元格区域中，返回特定行列交叉处单元格的值或引用',
  'function.INT': '将数值向下取整为最接近的整数',
  'function.ISBLANK': '检查是否引用了空单元格，返回 TRUE 或 FALSE',
  'function.ISERR': '检测一个值是否为 #N/A 以外的错误，返回 TRUE 或 FALSE',
  'function.ISERROR': '检测一个值是否为错误，返回 TRUE 或 FALSE',
  'function.ISEVEN': '如果数字为偶数则返回 TRUE',
  'function.ISFORMULA': '检查引用是否指向包含公式的单元格，返回 TRUE 或 FALSE',
  'function.ISLOGICAL':
    '检测一个值是否是逻辑值(TRUE 或 FALSE)，返回 TRUE 或 FALSE',
  'function.ISNA': '检测一个值是否为 #N/A，返回 TRUE 或 FALSE',
  'function.ISNONTEXT':
    '检测一个值是否不是文本(空白单元格不是文本)，返回 TRUE 或 FALSE',
  'function.ISNUMBER': '检测一个值是否是数值，返回 TRUE 或 FALSE',
  'function.ISODD': '如果数字为奇数则返回 TRUE',
  'function.ISOWEEKNUM': '返回给定日期在全年中的 ISO 周数',
  'function.ISREF': '检测一个值是否为引用，返回 TRUE 或 FALSE',
  'function.ISTEXT': '检测一个值是否为文本，返回 TRUE 或 FALSE',
  'function.LARGE': '返回数据组中第 k 个最大值。例如，第五个最大值',
  'function.LCM': '返回最小公倍数',
  'function.LEFT': '从一个文本字符串的第一个字符开始返回指定个数的字符',
  'function.LEN': '返回文本字符串中的字符个数',
  'function.LN': '返回给定数值的自然对数',
  'function.LOG': '根据给定底数返回数字的对数',
  'function.LOG10': '返回给定数值以 10 为底的对数',
  'function.LOOKUP': '从单行或单列或从数组中查找一个值。条件是向后兼容性',
  'function.LOWER': '将一个文本字符串的所有字母转换为小写形式',
  'function.MATCH': '返回符合特定值特定顺序的项在数组中的相对位置',
  'function.MAX': '返回一组数值中的最大值，忽略逻辑值及文本',
  'function.MAXA': '返回一组参数中的最大值(不忽略逻辑值和字符串)',
  'function.MAXIFS': '返回一组给定条件所指定的单元格的最大值',
  'function.MEDIAN': '返回一组数的中值',
  'function.MID': '从文本字符串中指定的起始位置起返回指定长度的字符',
  'function.MIN': '返回一组数值中的最小值，忽略逻辑值及文本',
  'function.MINA': '返回一组参数中的最小值(不忽略逻辑值和字符串)',
  'function.MINIFS': '返回一组给定条件所指定的单元格的最小值',
  'function.MINUTE': '返回分钟数值，是一个 0 到 59 之间的整数',
  'function.MOD': '返回两数相除的余数',
  'function.MODE': '返回一组数据或数据区域中出现频率最高或重复出现的数值',
  'function.MODE.SNGL': '返回一组数据或数据区域中出现频率最高或重复出现的数值',
  'function.MONTH': '返回月份值，是一个 1 (一月)到 12 (十二月)之间的数字',
  'function.MROUND': '返回一个舍入到所需倍数的数字',
  'function.N':
    '将不是数值形式的值转换为数值形式。日期转换成序列值，TRUE 转换成 1，其他值转换成 0',
  'function.NA': '返回错误值 #N/A (无法使用的数值)',
  'function.NETWORKDAYS': '返回两个日期之间的完整工作日数',
  'function.NETWORKDAYS.INTL':
    '使用自定义周末参数返回两个日期之间的完整工作日数',
  'function.NOT':
    '对参数的逻辑值求反：参数为 TRUE 时返回 FALSE；参数为 FALSE 时返回 TRUE',
  'function.NOW': '返回日期时间格式的当前日期和时间',
  'function.NUMBERVALUE': '按独立于区域设置的方式将文本转换为数字',
  'function.ODD': '将正(负)数向上(下)舍入到最接近的奇数',
  'function.OR':
    '如果任一参数值为 TRUE，即返回 TRUE；只有当所有参数值均为 FALSE 时才返回 FALSE',
  'function.PERCENTILE': '返回区域中数值的第 k 个百分点的值',
  'function.PERCENTILE.INC':
    '返回区域中数值的第 k 个百分点的值，k 介于 0 与 1 之间，含 0 与 1',
  'function.PI': '返回圆周率 Pi 的值，3.14159265358979，精确到 15 位',
  'function.POWER': '返回某数的乘幂',
  'function.PRODUCT': '计算所有参数的乘积',
  'function.PROPER':
    '将一个文本字符串中各英文单词的第一个字母转换成大写，将其他字母转换成小写',
  'function.QUARTILE': '返回一组数据的四分位点',
  'function.QUARTILE.INC':
    '基于从 0 到 1 之间(含 0 与 1)的百分点值，返回一组数据的四分位点',
  'function.QUOTIENT': '返回除法的整数部分',
  'function.RADIANS': '将角度转为弧度',
  'function.RANK': '返回某数字在一列数字中相对于其他数值的大小排名',
  'function.RANK.AVG':
    '返回某数字在一列数字中相对于其他数值的大小排名；如果多个数值排名相同，则返回平均值排名',
  'function.RANK.EQ':
    '返回某数字在一列数字中相对于其他数值的大小排名；如果多个数值排名相同，则返回该组数值的最佳排名',
  'function.REPLACE': '将一个字符串中的部分字符用另一个字符串替换',
  'function.REPT':
    '根据指定次数重复文本。可用 REPT 在一个单元格中重复填写一个文本字符串',
  'function.RIGHT': '从一个文本字符串的最后一个字符开始返回指定个数的字符',
  'function.ROUND': '按指定的位数对数值进行四舍五入',
  'function.ROUNDDOWN': '向下舍入数字',
  'function.ROUNDUP': '向上舍入数字',
  'function.ROW': '返回一个引用的行号',
  'function.ROWS': '返回某一引用或数组的行数',
  'function.SEARCH':
    '返回一个指定字符或文本字符串在字符串中第一次出现的位置(从左到右查找，忽略大小写)',
  'function.SECOND': '返回秒数值，是一个 0 到 59 之间的整数',
  'function.SIGN':
    '返回数字的正负号：为正时，返回 1；为零时，返回 0；为负时，返回 -1',
  'function.SIN': '返回给定角度的正弦值',
  'function.SINH': '返回双曲正弦值',
  'function.SMALL': '返回数据组中第 k 个最小值',
  'function.SQRT': '返回数值的平方根',
  'function.SQRTPI': '返回(数字 * Pi)的平方根',
  'function.STDEV': '估算基于给定样本的标准偏差(忽略样本中的逻辑值及文本)',
  'function.STDEV.P': '计算基于给定的样本总体的标准偏差(忽略逻辑值及文本)',
  'function.STDEV.S': '估算基于样本的标准偏差(忽略样本中的逻辑值及文本)',
  'function.STDEVP': '计算基于给定的样本总体的标准偏差(忽略逻辑值及文本)',
  'function.SUBSTITUTE': '将字符串中的部分字符串以新字符串替换',
  'function.SUM': '计算单元格区域中所有数值的和',
  'function.SUMIF': '对满足条件的单元格求和',
  'function.SUMIFS': '对一组给定条件指定的单元格求和',
  'function.SUMPRODUCT': '返回相应的数组或区域乘积的和',
  'function.SUMSQ':
    '返回所有参数的平方和。参数可以是数值、数组、名称，或者是对数值单元格的引用',
  'function.SWITCH':
    '根据值列表计算表达式并返回与第一个匹配值对应的结果。如果没有匹配项，则返回可选默认值',
  'function.T':
    '检测给定值是否为文本，如果是文本按原样返回，如果不是文本则返回双引号(空文本)',
  'function.TAN': '返回给定角度的正切值',
  'function.TANH': '返回双曲正切值',
  'function.TEXT': '根据指定的数值格式将数字转成文本',
  'function.TEXTJOIN': '使用分隔符连接列表或文本字符串区域',
  'function.TIME': '返回特定时间的序列数',
  'function.TIMEVALUE':
    '将文本形式表示的时间转换成 Excel 序列数，是一个从 0 (12:00:00 AM) 到 0.999988426 (11:59:59 PM) 的数',
  'function.TODAY': '返回日期格式的当前日期',
  'function.TRIM':
    '删除字符串中多余的空格，但会在英文字符串中保留一个作为词与词之间分隔的空格',
  'function.TRUE': '返回逻辑值 TRUE',
  'function.TRUNC': '将数字截为整数或保留指定位数的小数',
  'function.TYPE':
    '以整数形式返回值的数据类型：数值 = 1；文字 = 2；逻辑值 = 4；错误值 = 16；数组 = 64',
  'function.UNICHAR': '返回由给定数值引用的 Unicode 字符',
  'function.UNICODE': '返回对应于文本的第一个字符的数字(代码点)',
  'function.UPPER': '将文本字符串转换成字母全部大写形式',
  'function.VALUE': '将一个代表数值的文本字符串转换成数值',
  'function.VAR': '估算基于给定样本的方差(忽略样本中的逻辑值及文本)',
  'function.VAR.P': '计算基于整个样本总体的方差(忽略样本总体中的逻辑值及文本)',
  'function.VAR.S': '估算基于给定样本的方差(忽略样本中的逻辑值及文本)',
  'function.VARP': '计算基于样本总体的方差(忽略样本总体中的逻辑值及文本)',
  'function.VLOOKUP':
    '搜索表区域首列满足条件的元素，确定待检索单元格在区域中的行序号，再进一步返回选定单元格的值。默认情况下，表是以升序排序的',
  'function.WEEKDAY': '返回代表一周中的第几天的数值，是一个 1 到 7 之间的整数',
  'function.WEEKNUM': '返回一年中的周数',
  'function.WORKDAY': '返回在指定的若干个工作日之前/之后的日期(一串数字)',
  'function.WORKDAY.INTL':
    '使用自定义周末参数返回在指定的若干个工作日之前/之后的日期(一串数字)',
  'function.XLOOKUP':
    '搜索区域或数组以查找匹配项，并从第二个区域或数组返回相应的项。默认情况下使用精确匹配',
  'function.XMATCH': '返回项在数组中的相对位置。默认情况下，需要精确匹配',
  'function.XOR': '返回所有参数的逻辑“异或”值',
  'function.YEAR': '返回日期的年份值，一个 1900-9999 之间的数字',
  'function.YEARFRAC':
    '返回一个年分数，表示 start_date 和 end_date 之间的整天天数',
  'value.ADDRESS.abs_num.1': '绝对引用',
  'value.ADDRESS.abs_num.2': '绝对行；相对列',
  'value.ADDRESS.abs_num.3': '相对行；绝对列',
  'value.ADDRESS.abs_num.4': '相对引用',
  'value.ADDRESS.a1.TRUE': 'A1 样式',
  'value.ADDRESS.a1.FALSE': 'R1C1 样式',
  'value.DAYS360.method.FALSE': '美国(NASD)方法',
  'value.DAYS360.method.TRUE': '欧洲方法',
  'value.HLOOKUP.range_lookup.TRUE': '近似匹配',
  'value.HLOOKUP.range_lookup.FALSE': '精确匹配',
  'value.MATCH.match_type.1': '小于',
  'value.MATCH.match_type.0': '精确匹配',
  'value.MATCH.match_type.-1': '大于',
  'value.NETWORKDAYS.INTL.weekend.1': '星期六、星期日',
  'value.NETWORKDAYS.INTL.weekend.2': '星期日、星期一',
  'value.NETWORKDAYS.INTL.weekend.3': '星期一、星期二',
  'value.NETWORKDAYS.INTL.weekend.4': '星期二、星期三',
  'value.NETWORKDAYS.INTL.weekend.5': '星期三、星期四',
  'value.NETWORKDAYS.INTL.weekend.6': '星期四、星期五',
  'value.NETWORKDAYS.INTL.weekend.7': '星期五、星期六',
  'value.NETWORKDAYS.INTL.weekend.11': '仅星期日',
  'value.NETWORKDAYS.INTL.weekend.12': '仅星期一',
  'value.NETWORKDAYS.INTL.weekend.13': '仅星期二',
  'value.NETWORKDAYS.INTL.weekend.14': '仅星期三',
  'value.NETWORKDAYS.INTL.weekend.15': '仅星期四',
  'value.NETWORKDAYS.INTL.weekend.16': '仅星期五',
  'value.NETWORKDAYS.INTL.weekend.17': '仅星期六',
  'value.RANK.order.0': '降序',
  'value.RANK.order.1': '升序',
  'value.RANK.AVG.order.0': '降序',
  'value.RANK.AVG.order.1': '升序',
  'value.RANK.EQ.order.0': '降序',
  'value.RANK.EQ.order.1': '升序',
  'value.VLOOKUP.range_lookup.TRUE': '近似匹配',
  'value.VLOOKUP.range_lookup.FALSE': '精确匹配',
  'value.WEEKDAY.return_type.1': '从 1 (星期日)到 7 (星期六)的数字',
  'value.WEEKDAY.return_type.2': '从 1 (星期一)到 7 (星期日)的数字',
  'value.WEEKDAY.return_type.3': '从 0 (星期一)到 6 (星期日)的数字',
  'value.WEEKDAY.return_type.11': '从 1 (星期一)到 7 (星期日)的数字',
  'value.WEEKDAY.return_type.12': '从 1 (星期二)到 7 (星期一)的数字',
  'value.WEEKDAY.return_type.13': '从 1 (星期三)到 7 (星期二)的数字',
  'value.WEEKDAY.return_type.14': '从 1 (星期四)到 7 (星期三)的数字',
  'value.WEEKDAY.return_type.15': '从 1 (星期五)到 7 (星期四)的数字',
  'value.WEEKDAY.return_type.16': '从 1 (星期六)到 7 (星期五)的数字',
  'value.WEEKDAY.return_type.17': '从 1 (星期日)到 7 (星期六)的数字',
  'value.WEEKNUM.return_type.1': '一周从星期日开始',
  'value.WEEKNUM.return_type.2': '一周从星期一开始',
  'value.WEEKNUM.return_type.11': '一周从星期一开始',
  'value.WEEKNUM.return_type.12': '一周从星期二开始',
  'value.WEEKNUM.return_type.13': '一周从星期三开始',
  'value.WEEKNUM.return_type.14': '一周从星期四开始',
  'value.WEEKNUM.return_type.15': '一周从星期五开始',
  'value.WEEKNUM.return_type.16': '一周从星期六开始',
  'value.WEEKNUM.return_type.17': '一周从星期日开始',
  'value.WEEKNUM.return_type.21': '一周从星期一开始(系统 2)',
  'value.WORKDAY.INTL.weekend.1': '星期六、星期日',
  'value.WORKDAY.INTL.weekend.2': '星期日、星期一',
  'value.WORKDAY.INTL.weekend.3': '星期一、星期二',
  'value.WORKDAY.INTL.weekend.4': '星期二、星期三',
  'value.WORKDAY.INTL.weekend.5': '星期三、星期四',
  'value.WORKDAY.INTL.weekend.6': '星期四、星期五',
  'value.WORKDAY.INTL.weekend.7': '星期五、星期六',
  'value.WORKDAY.INTL.weekend.11': '仅星期日',
  'value.WORKDAY.INTL.weekend.12': '仅星期一',
  'value.WORKDAY.INTL.weekend.13': '仅星期二',
  'value.WORKDAY.INTL.weekend.14': '仅星期三',
  'value.WORKDAY.INTL.weekend.15': '仅星期四',
  'value.WORKDAY.INTL.weekend.16': '仅星期五',
  'value.WORKDAY.INTL.weekend.17': '仅星期六',
  'value.XLOOKUP.match_mode.0': '完全匹配',
  'value.XLOOKUP.match_mode.-1': '完全匹配或下一个较小的项',
  'value.XLOOKUP.match_mode.1': '完全匹配或下一个较大的项',
  'value.XLOOKUP.match_mode.2': '通配符匹配',
  'value.XLOOKUP.search_mode.1': '从第一项到最后一项搜索',
  'value.XLOOKUP.search_mode.-1': '从最后一项到第一项搜索',
  'value.XLOOKUP.search_mode.2': '二进制搜索(按升序排序)',
  'value.XLOOKUP.search_mode.-2': '二进制搜索(按降序排序)',
  'value.XMATCH.match_mode.0': '完全匹配',
  'value.XMATCH.match_mode.-1': '完全匹配或下一个较小的项',
  'value.XMATCH.match_mode.1': '完全匹配或下一个较大的项',
  'value.XMATCH.match_mode.2': '通配符匹配',
  'value.XMATCH.search_mode.1': '从第一项到最后一项搜索',
  'value.XMATCH.search_mode.-1': '从最后一项到第一项搜索',
  'value.XMATCH.search_mode.2': '二进制搜索(按升序排序)',
  'value.XMATCH.search_mode.-2': '二进制搜索(按降序排序)',
  'value.YEARFRAC.basis.0': '美国(NASD) 30/360',
  'value.YEARFRAC.basis.1': '实际天数/实际天数',
  'value.YEARFRAC.basis.2': '实际天数/360',
  'value.YEARFRAC.basis.3': '实际天数/365',
  'value.YEARFRAC.basis.4': '欧洲 30/360',
};
