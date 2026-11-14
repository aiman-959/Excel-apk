/** Argument names shown under the formula bar while typing a function. */
const RAW = `
SUM:number1,[number2],…|AVERAGE:number1,[number2],…|COUNT:value1,[value2],…|COUNTA:value1,[value2],…|COUNTBLANK:range|MAX:number1,[number2],…|MIN:number1,[number2],…
MEDIAN:number1,[number2],…|MODE:number1,[number2],…|PRODUCT:number1,[number2],…|STDEV:number1,[number2],…|VAR:number1,[number2],…|LARGE:array,k|SMALL:array,k|RANK:number,ref,[order]
IF:logical_test,value_if_true,[value_if_false]|IFS:logical_test1,value_if_true1,[logical_test2,value_if_true2],…|IFERROR:value,value_if_error|IFNA:value,value_if_na
AND:logical1,[logical2],…|OR:logical1,[logical2],…|NOT:logical|XOR:logical1,[logical2],…|SWITCH:expression,value1,result1,[default_or_value2,result2],…
VLOOKUP:lookup_value,table_array,col_index_num,[range_lookup]|HLOOKUP:lookup_value,table_array,row_index_num,[range_lookup]|XLOOKUP:lookup_value,lookup_array,return_array,[if_not_found],[match_mode],[search_mode]
INDEX:array,row_num,[column_num]|MATCH:lookup_value,lookup_array,[match_type]|XMATCH:lookup_value,lookup_array,[match_mode],[search_mode]|LOOKUP:lookup_value,lookup_vector,[result_vector]
CHOOSE:index_num,value1,[value2],…|OFFSET:reference,rows,cols,[height],[width]|INDIRECT:ref_text,[a1]|ROW:[reference]|COLUMN:[reference]|ROWS:array|COLUMNS:array|ADDRESS:row_num,column_num,[abs_num],[a1],[sheet_text]
SUMIF:range,criteria,[sum_range]|SUMIFS:sum_range,criteria_range1,criteria1,[criteria_range2,criteria2],…|COUNTIF:range,criteria|COUNTIFS:criteria_range1,criteria1,[criteria_range2,criteria2],…
AVERAGEIF:range,criteria,[average_range]|AVERAGEIFS:average_range,criteria_range1,criteria1,…|MAXIFS:max_range,criteria_range1,criteria1,…|MINIFS:min_range,criteria_range1,criteria1,…|SUMPRODUCT:array1,[array2],…
ROUND:number,num_digits|ROUNDUP:number,num_digits|ROUNDDOWN:number,num_digits|INT:number|TRUNC:number,[num_digits]|CEILING:number,significance|FLOOR:number,significance|MROUND:number,multiple
ABS:number|SQRT:number|POWER:number,power|MOD:number,divisor|QUOTIENT:numerator,denominator|EXP:number|LN:number|LOG:number,[base]|LOG10:number|PI:|RAND:|RANDBETWEEN:bottom,top|SIGN:number
TEXT:value,format_text|VALUE:text|CONCAT:text1,[text2],…|CONCATENATE:text1,[text2],…|TEXTJOIN:delimiter,ignore_empty,text1,[text2],…|LEFT:text,[num_chars]|RIGHT:text,[num_chars]|MID:text,start_num,num_chars
LEN:text|TRIM:text|UPPER:text|LOWER:text|PROPER:text|SUBSTITUTE:text,old_text,new_text,[instance_num]|REPLACE:old_text,start_num,num_chars,new_text|FIND:find_text,within_text,[start_num]|SEARCH:find_text,within_text,[start_num]
REPT:text,number_times|EXACT:text1,text2|CHAR:number|CODE:text|CLEAN:text|TEXTBEFORE:text,delimiter,[instance_num]|TEXTAFTER:text,delimiter,[instance_num]|TEXTSPLIT:text,col_delimiter,[row_delimiter]
DATE:year,month,day|TODAY:|NOW:|YEAR:serial_number|MONTH:serial_number|DAY:serial_number|HOUR:serial_number|MINUTE:serial_number|SECOND:serial_number|TIME:hour,minute,second
WEEKDAY:serial_number,[return_type]|WEEKNUM:serial_number,[return_type]|EOMONTH:start_date,months|EDATE:start_date,months|DATEDIF:start_date,end_date,unit|DAYS:end_date,start_date
NETWORKDAYS:start_date,end_date,[holidays]|WORKDAY:start_date,days,[holidays]|DATEVALUE:date_text|TIMEVALUE:time_text|YEARFRAC:start_date,end_date,[basis]
ISBLANK:value|ISNUMBER:value|ISTEXT:value|ISERROR:value|ISNA:value|ISLOGICAL:value|ISEVEN:number|ISODD:number|N:value|NA:|TYPE:value
PMT:rate,nper,pv,[fv],[type]|PV:rate,nper,pmt,[fv],[type]|FV:rate,nper,pmt,[pv],[type]|NPV:rate,value1,[value2],…|IRR:values,[guess]|RATE:nper,pmt,pv,[fv],[type],[guess]|NPER:rate,pmt,pv,[fv],[type]
FILTER:array,include,[if_empty]|SORT:array,[sort_index],[sort_order],[by_col]|UNIQUE:array,[by_col],[exactly_once]|SEQUENCE:rows,[columns],[start],[step]|TRANSPOSE:array|SORTBY:array,by_array1,[sort_order1],…
TAKE:array,rows,[columns]|DROP:array,rows,[columns]|CHOOSECOLS:array,col_num1,[col_num2],…|CHOOSEROWS:array,row_num1,[row_num2],…|VSTACK:array1,[array2],…|HSTACK:array1,[array2],…|TOCOL:array,[ignore],[scan_by_column]|TOROW:array,[ignore],[scan_by_column]
HYPERLINK:link_location,[friendly_name]|SUBTOTAL:function_num,ref1,[ref2],…|AGGREGATE:function_num,options,ref1,…|CORREL:array1,array2|FORECAST:x,known_ys,known_xs|SLOPE:known_ys,known_xs|INTERCEPT:known_ys,known_xs
PERCENTILE:array,k|QUARTILE:array,quart|AVERAGEA:value1,[value2],…|STDEV.S:number1,[number2],…|STDEV.P:number1,[number2],…|VAR.S:number1,[number2],…|VAR.P:number1,[number2],…|GEOMEAN:number1,[number2],…
CONVERT:number,from_unit,to_unit|DEC2BIN:number,[places]|BIN2DEC:number|DEC2HEX:number,[places]|HEX2DEC:number|GCD:number1,[number2],…|LCM:number1,[number2],…|FACT:number|COMBIN:number,number_chosen
SIN:number|COS:number|TAN:number|ASIN:number|ACOS:number|ATAN:number|ATAN2:x_num,y_num|RADIANS:angle|DEGREES:angle|LET:name1,name_value1,calculation_or_name2,…|NUMBERVALUE:text,[decimal_separator],[group_separator]
`;
export const FUNCTION_HINTS: Record<string, string[]> = {};
for (const item of RAW.split(/[|\n]/)) {
  const i = item.indexOf(':');
  if (i > 0) FUNCTION_HINTS[item.slice(0, i).trim()] = item.slice(i + 1).split(',').filter(Boolean);
}
