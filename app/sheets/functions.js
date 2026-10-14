/* BTS Sheets — function catalogue: name | category | arguments | description */
(function () {
  const RAW = `
SUM|Math|number1, [number2], ...|Adds all the numbers in a range of cells.
SUMIF|Math|range, criteria, [sum_range]|Adds the cells that meet a condition.
SUMIFS|Math|sum_range, criteria_range1, criteria1, ...|Adds the cells that meet several conditions.
SUMPRODUCT|Math|array1, [array2], ...|Returns the sum of the products of matching ranges.
SUMSQ|Math|number1, [number2], ...|Returns the sum of the squares of the arguments.
PRODUCT|Math|number1, [number2], ...|Multiplies all the numbers given as arguments.
ROUND|Math|number, num_digits|Rounds a number to a specified number of digits.
ROUNDUP|Math|number, num_digits|Rounds a number up, away from zero.
ROUNDDOWN|Math|number, num_digits|Rounds a number down, toward zero.
MROUND|Math|number, multiple|Rounds a number to the nearest multiple.
CEILING|Math|number, significance|Rounds a number up to the nearest multiple of significance.
FLOOR|Math|number, significance|Rounds a number down to the nearest multiple of significance.
INT|Math|number|Rounds a number down to the nearest integer.
TRUNC|Math|number, [num_digits]|Truncates a number to an integer.
ABS|Math|number|Returns the absolute value of a number.
MOD|Math|number, divisor|Returns the remainder after division.
QUOTIENT|Math|numerator, denominator|Returns the integer portion of a division.
POWER|Math|number, power|Returns the result of a number raised to a power.
SQRT|Math|number|Returns a positive square root.
EXP|Math|number|Returns e raised to the power of a number.
LN|Math|number|Returns the natural logarithm of a number.
LOG|Math|number, [base]|Returns the logarithm of a number to the base you specify.
LOG10|Math|number|Returns the base-10 logarithm of a number.
PI|Math||Returns the value of pi.
RAND|Math||Returns a random number between 0 and 1.
RANDBETWEEN|Math|bottom, top|Returns a random integer between the numbers you specify.
SIGN|Math|number|Returns the sign of a number.
EVEN|Math|number|Rounds a number up to the nearest even integer.
ODD|Math|number|Rounds a number up to the nearest odd integer.
FACT|Math|number|Returns the factorial of a number.
COMBIN|Math|number, number_chosen|Returns the number of combinations.
GCD|Math|number1, [number2], ...|Returns the greatest common divisor.
LCM|Math|number1, [number2], ...|Returns the least common multiple.
SIN|Math|number|Returns the sine of an angle (radians).
COS|Math|number|Returns the cosine of an angle (radians).
TAN|Math|number|Returns the tangent of an angle (radians).
RADIANS|Math|angle|Converts degrees to radians.
DEGREES|Math|angle|Converts radians to degrees.
SUBTOTAL|Math|function_num, ref1, ...|Returns a subtotal in a list.
ROMAN|Math|number, [form]|Converts an Arabic numeral to Roman, as text.
AVERAGE|Statistical|number1, [number2], ...|Returns the average (arithmetic mean) of its arguments.
AVERAGEA|Statistical|value1, [value2], ...|Average including text and logical values.
AVERAGEIF|Statistical|range, criteria, [average_range]|Average of cells that meet a condition.
AVERAGEIFS|Statistical|average_range, criteria_range1, criteria1, ...|Average of cells that meet several conditions.
COUNT|Statistical|value1, [value2], ...|Counts the cells that contain numbers.
COUNTA|Statistical|value1, [value2], ...|Counts the cells that are not empty.
COUNTBLANK|Statistical|range|Counts empty cells in a range.
COUNTIF|Statistical|range, criteria|Counts the cells that meet a condition.
COUNTIFS|Statistical|criteria_range1, criteria1, ...|Counts the cells that meet several conditions.
MAX|Statistical|number1, [number2], ...|Returns the largest value.
MIN|Statistical|number1, [number2], ...|Returns the smallest value.
MAXIFS|Statistical|max_range, criteria_range1, criteria1, ...|Largest value among cells that meet conditions.
MINIFS|Statistical|min_range, criteria_range1, criteria1, ...|Smallest value among cells that meet conditions.
MEDIAN|Statistical|number1, [number2], ...|Returns the median of the given numbers.
MODE.SNGL|Statistical|number1, [number2], ...|Returns the most frequently occurring value.
LARGE|Statistical|array, k|Returns the k-th largest value.
SMALL|Statistical|array, k|Returns the k-th smallest value.
RANK.EQ|Statistical|number, ref, [order]|Returns the rank of a number in a list.
RANK.AVG|Statistical|number, ref, [order]|Returns the rank of a number, averaging ties.
STDEV.S|Statistical|number1, [number2], ...|Estimates standard deviation based on a sample.
STDEV.P|Statistical|number1, [number2], ...|Standard deviation of an entire population.
STDEV|Statistical|number1, [number2], ...|Estimates standard deviation based on a sample.
VAR.S|Statistical|number1, [number2], ...|Estimates variance based on a sample.
VAR.P|Statistical|number1, [number2], ...|Variance of an entire population.
PERCENTILE.INC|Statistical|array, k|Returns the k-th percentile of values.
QUARTILE.INC|Statistical|array, quart|Returns the quartile of a data set.
CORREL|Statistical|array1, array2|Returns the correlation coefficient.
SLOPE|Statistical|known_ys, known_xs|Returns the slope of the linear regression line.
INTERCEPT|Statistical|known_ys, known_xs|Returns the intercept of the linear regression line.
FORECAST.LINEAR|Statistical|x, known_ys, known_xs|Predicts a value along a linear trend.
GEOMEAN|Statistical|number1, [number2], ...|Returns the geometric mean.
IF|Logical|logical_test, [value_if_true], [value_if_false]|Returns one value if a condition is TRUE and another if FALSE.
IFS|Logical|logical_test1, value_if_true1, ...|Checks conditions and returns the value for the first TRUE one.
IFERROR|Logical|value, value_if_error|Returns value_if_error if the expression is an error.
IFNA|Logical|value, value_if_na|Returns value_if_na if the expression is #N/A.
AND|Logical|logical1, [logical2], ...|TRUE if all arguments are TRUE.
OR|Logical|logical1, [logical2], ...|TRUE if any argument is TRUE.
NOT|Logical|logical|Reverses the logic of its argument.
XOR|Logical|logical1, [logical2], ...|Exclusive OR of all arguments.
SWITCH|Logical|expression, value1, result1, [default]|Matches an expression against a list of values.
TRUE|Logical||Returns the logical value TRUE.
FALSE|Logical||Returns the logical value FALSE.
CONCATENATE|Text|text1, [text2], ...|Joins several text items into one.
TEXTJOIN|Text|delimiter, ignore_empty, text1, ...|Joins text with a delimiter.
LEFT|Text|text, [num_chars]|Returns the leftmost characters of a text value.
RIGHT|Text|text, [num_chars]|Returns the rightmost characters of a text value.
MID|Text|text, start_num, num_chars|Returns characters from the middle of a text string.
LEN|Text|text|Returns the number of characters in a text string.
LOWER|Text|text|Converts text to lowercase.
UPPER|Text|text|Converts text to uppercase.
PROPER|Text|text|Capitalises the first letter of each word.
TRIM|Text|text|Removes extra spaces from text.
CLEAN|Text|text|Removes non-printable characters.
SUBSTITUTE|Text|text, old_text, new_text, [instance_num]|Substitutes new text for old text.
REPLACE|Text|old_text, start_num, num_chars, new_text|Replaces characters within text.
FIND|Text|find_text, within_text, [start_num]|Finds one text value within another (case-sensitive).
SEARCH|Text|find_text, within_text, [start_num]|Finds one text value within another (not case-sensitive).
TEXT|Text|value, format_text|Formats a number and converts it to text.
REPT|Text|text, number_times|Repeats text a given number of times.
EXACT|Text|text1, text2|Checks whether two text values are identical.
CHAR|Text|number|Returns the character specified by the code number.
CODE|Text|text|Returns a numeric code for the first character.
UNICHAR|Text|number|Returns the Unicode character for a number.
T|Text|value|Returns the text referred to by value.
N|Text|value|Converts a value to a number.
DATE|Date & Time|year, month, day|Returns the serial number of a particular date.
DATEVALUE|Date & Time|date_text|Converts a date in text form to a serial number.
TIME|Date & Time|hour, minute, second|Returns the serial number of a particular time.
TIMEVALUE|Date & Time|time_text|Converts a time in text form to a serial number.
TODAY|Date & Time||Returns today's date.
NOW|Date & Time||Returns the current date and time.
YEAR|Date & Time|serial_number|Returns the year of a date.
MONTH|Date & Time|serial_number|Returns the month of a date (1–12).
DAY|Date & Time|serial_number|Returns the day of the month (1–31).
HOUR|Date & Time|serial_number|Returns the hour of a time value.
MINUTE|Date & Time|serial_number|Returns the minute of a time value.
SECOND|Date & Time|serial_number|Returns the second of a time value.
WEEKDAY|Date & Time|serial_number, [return_type]|Returns the day of the week.
WEEKNUM|Date & Time|serial_number, [return_type]|Returns the week number of the year.
EDATE|Date & Time|start_date, months|Date that is a number of months before or after a date.
EOMONTH|Date & Time|start_date, months|Last day of the month before or after a number of months.
DATEDIF|Date & Time|start_date, end_date, unit|Difference between two dates in days, months or years.
DAYS|Date & Time|end_date, start_date|Number of days between two dates.
DAYS360|Date & Time|start_date, end_date, [method]|Days between two dates based on a 360-day year.
NETWORKDAYS|Date & Time|start_date, end_date, [holidays]|Number of whole working days between two dates.
WORKDAY|Date & Time|start_date, days, [holidays]|Date a number of working days away.
YEARFRAC|Date & Time|start_date, end_date, [basis]|Fraction of the year between two dates.
VLOOKUP|Lookup|lookup_value, table_array, col_index_num, [range_lookup]|Looks in the first column of a table and returns a value in the same row.
HLOOKUP|Lookup|lookup_value, table_array, row_index_num, [range_lookup]|Looks in the top row of a table and returns a value in the same column.
XLOOKUP|Lookup|lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]|Searches a range and returns the matching item.
LOOKUP|Lookup|lookup_value, lookup_vector, [result_vector]|Looks up a value in a one-row or one-column range.
INDEX|Lookup|array, row_num, [column_num]|Returns a value from a table by row and column number.
MATCH|Lookup|lookup_value, lookup_array, [match_type]|Returns the position of an item in a range.
CHOOSE|Lookup|index_num, value1, [value2], ...|Chooses a value from a list by index.
ROW|Lookup|[reference]|Returns the row number of a reference.
ROWS|Lookup|array|Returns the number of rows in a reference.
COLUMN|Lookup|[reference]|Returns the column number of a reference.
COLUMNS|Lookup|array|Returns the number of columns in a reference.
TRANSPOSE|Lookup|array|Returns the transpose of an array.
HYPERLINK|Lookup|link_location, [friendly_name]|Creates a link that opens a web page or file.
ADDRESS|Lookup|row_num, column_num, [abs_num]|Returns a cell reference as text.
FORMULATEXT|Lookup|reference|Returns the formula of a cell as text.
PMT|Financial|rate, nper, pv, [fv], [type]|Payment for a loan based on constant payments and interest rate.
IPMT|Financial|rate, per, nper, pv, [fv], [type]|Interest payment for a given period.
PPMT|Financial|rate, per, nper, pv, [fv], [type]|Principal payment for a given period.
FV|Financial|rate, nper, pmt, [pv], [type]|Future value of an investment.
PV|Financial|rate, nper, pmt, [fv], [type]|Present value of an investment.
NPV|Financial|rate, value1, [value2], ...|Net present value of cash flows.
NPER|Financial|rate, pmt, pv, [fv], [type]|Number of periods for an investment.
RATE|Financial|nper, pmt, pv, [fv], [type], [guess]|Interest rate per period of an annuity.
EFFECT|Financial|nominal_rate, npery|Effective annual interest rate.
NOMINAL|Financial|effect_rate, npery|Nominal annual interest rate.
SLN|Financial|cost, salvage, life|Straight-line depreciation for one period.
DB|Financial|cost, salvage, life, period, [month]|Fixed-declining balance depreciation.
DDB|Financial|cost, salvage, life, period, [factor]|Double-declining balance depreciation.
SYD|Financial|cost, salvage, life, per|Sum-of-years' digits depreciation.
CUMIPMT|Financial|rate, nper, pv, start_period, end_period, type|Cumulative interest paid between two periods.
CUMPRINC|Financial|rate, nper, pv, start_period, end_period, type|Cumulative principal paid between two periods.
ISBLANK|Information|value|TRUE if the value is blank.
ISNUMBER|Information|value|TRUE if the value is a number.
ISTEXT|Information|value|TRUE if the value is text.
ISERROR|Information|value|TRUE if the value is any error.
ISNA|Information|value|TRUE if the value is #N/A.
ISEVEN|Information|number|TRUE if the number is even.
ISODD|Information|number|TRUE if the number is odd.
ISLOGICAL|Information|value|TRUE if the value is a logical value.
ISFORMULA|Information|reference|TRUE if the cell contains a formula.
NA|Information||Returns the error value #N/A.
`;
  const list = RAW.trim().split('\n').map(l => { const [name, cat, args, desc] = l.split('|'); return { name, cat, args, desc }; });
  const byName = new Map(list.map(f => [f.name, f]));
  window.SFUNCS = { list, byName, cats: ['Financial', 'Logical', 'Text', 'Date & Time', 'Lookup', 'Math', 'Statistical', 'Information'] };
})();
