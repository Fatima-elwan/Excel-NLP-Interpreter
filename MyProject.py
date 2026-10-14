import re
import openpyxl
import pandas as pd 
import matplotlib.pyplot as plt
from collections import Counter
import math

file_path = "students.xlsx"   
sheet_name = "students"

try:
    df = pd.read_excel(file_path)
except Exception:
    df = None


def tokenize(expression):

    expression = expression.replace(" ", "")
    
    
    tokens = re.findall(r'=[A-Za-z]+|[A-Za-z][A-Za-z0-9]*!?[A-Za-z]+\d+|[A-Za-z]+\d+|\d+|\+|>=|<=|>|<|=|-|\*|/|\(|\)|\:|;|[A-Za-z]+|=|,', expression)
    # tokens = [t.strip() for t in tokens if t and t.strip()]
    if tokens and tokens[0].startswith('='):
        first = tokens.pop(0)
        if len(first) > 1:
            tokens.insert(0, first[1:])
            tokens.insert(0, '=')
        else:
            tokens.insert(0, '=')
            
    return tokens

def is_valid_cell_reference(cell_ref):
    valid_pattern = r'^([A-Za-z][A-Za-z0-9]*!)?[A-Za-z]+\d+$'
    if re.match(valid_pattern, cell_ref):
        return True
    return False 

def parse_cell_reference(cell_ref):
    if '!' in cell_ref:
        sheet, cell = cell_ref.split('!', 1)
        return sheet, cell
    return None, cell_ref 


def math_expression(sentence):
    smaller = sentence.lower().strip()
    
    if "square root of" in smaller:
        smaller = smaller.replace("square root of", "sqrt")
    
    if "evaluate the following expression:" in smaller:
        clean_math = smaller.replace("evaluate the following expression:", "")
    elif "evaluate the expression:" in smaller:
        clean_math = smaller.replace("evaluate the expression:", "")
    else:
        clean_math = smaller.replace("evaluate", "")
    clean_math = clean_math.strip()
    if "evaluate" in smaller and clean_math:
        return clean_math

    if "starting from" not in smaller and "from the" not in smaller and "if " not in smaller:
        if "mark" in smaller or "score" in smaller or "grade" in smaller or "student" in smaller:
            ordinal_to_row = {
                "first": "D2", "second": "D3", "third": "D4", "fourth": "D5", "fifth": "D6",
                "sixth": "D7", "seventh": "D8", "eighth": "D9", "ninth": "D10", "tenth": "D11"
            }
            found_words = [word for word in ordinal_to_row if word in smaller]
            if len(found_words) > 1 or ("and" in smaller and len(found_words) == 1):
                found_cells = [ordinal_to_row[w] for w in found_words]
                if "sum" in smaller or "total" in smaller or "add" in smaller:
                    return " + ".join(found_cells)
                elif "average" in smaller or "mean" in smaller:
                    return f"({ ' + '.join(found_cells) }) / {len(found_cells)}"

    if smaller.startswith("=") or any(x in smaller for x in ['sum', 'average', 'max', 'min', 'count']):
        return None
        
    math_exp = re.search(r"(sqrt\s*\(\s*[A-Za-z0-9_+\-*/^().\s]+\s*\)|sqrt\s+[a-zA-z0-9]+|\(?\d+[\d\s\+\-\*\/\^\(\)]+[\d\s\+\-\*\/\^\(\)]*\d+\)?|[A-Za-f]\d+\s*[\+\-\*\/]\s*[A-Za-f]\d+)", smaller)
    if math_exp:
        expression = math_exp.group().strip()
        if any(char in expression for char in "+-*/^()") or "sqrt" in expression:
            return expression
            
    return None

def natural(sentence):
    smaller = sentence.lower().strip()
    
    smaller = smaller.replace("greater than or equal to", ">=")
    smaller = smaller.replace("less than or equal to", "<=")
    smaller = smaller.replace("greater than", ">").replace("more than", ">").replace("above", ">")
    smaller = smaller.replace("less than", "<").replace("under", "<").replace("below", "<")
    smaller = smaller.replace("equal to", "=").replace(" equals ", "=").replace(" is ", "=")

    operation = ""
    return_type = "Value"
    end_row = len(df) + 1 if df is not None else 11
    
   
    def column(smaller_text):
        if re.search(r"\b(mark|score|grade|scores|marks|grades)\b", smaller_text): return "D"
        elif "weight" in smaller_text: return "E"
        elif "height" in smaller_text: return "F"
        elif re.search(r"\b(name|names)\b", smaller_text): return "A"
        elif re.search(r"\bage\b", smaller_text): return "B"
        elif re.search(r"\b(gender|sex|female|male|females|males)\b", smaller_text): return "C"
        elif re.search(r"\b[fm]\b", smaller_text) or smaller_text.strip() in ['f', 'm']: return "C"
        
        cell_match = re.search(r'\b([a-f])\d+\b', smaller_text)
        if cell_match:
            return cell_match.group(1).upper()
        return ""

    ordinal_rows = {
        "first": 2, "second": 3, "third": 4, "fourth": 5, "fifth": 6,
        "sixth": 7, "seventh": 8, "eighth": 9, "ninth": 10, "tenth": 11
    }
    
    start_row = 2 
    found_ordinal = False
    for word, row_num in ordinal_rows.items():
        if word in smaller:
            start_row = row_num
            found_ordinal = True
            break
            
    if not found_ordinal:
        row_match = re.search(r"row\s*(\d+)", smaller)
        if row_match:
            start_row = int(row_match.group(1))

    def parse_single_condition(text):
        text_clean = text.lower().replace("check if", "").replace("if", "").strip()
        text_clean = re.sub(r"\bis\b", "=", text_clean)
        
        col = column(text_clean)
        if not col: 
            col = column(smaller)
            if not col: return None
        
        operator = ""
        if ">=" in text_clean: operator = ">="
        elif "<=" in text_clean: operator = "<="
        elif ">" in text_clean: operator = ">"
        elif "<" in text_clean: operator = "<"
        elif "=" in text_clean: operator = "="
        
        if operator == "": return None
        
        parts = text_clean.split(operator)
        if len(parts) < 2: return None

                
        words_after = parts[-1].strip().split()
        if not words_after: return None
        raw_target = words_after[0].replace('"', '').replace("'", "")
        
        if col == "C" and raw_target in ['m', 'f']:
            raw_target = f"'{raw_target}'"

            
        return col, operator, raw_target

    true_output = "Pass"
    false_output = "Fail"
    out_block = smaller.strip()
    success_match = re.search(r"(?:then|output|give|return)\s+[\"']?([a-zA-Z0-9]+)[\"']?", out_block, re.IGNORECASE)
    if success_match: true_output = success_match.group(1).capitalize()
    fail_match = re.search(r"(?:else|otherwise)\s+[\"']?([a-zA-Z0-9]+)[\"']?", out_block, re.IGNORECASE)
    if fail_match: false_output = fail_match.group(1).capitalize()

    if re.search(r"\bwho\b|\bnames?\b|\bwhich student\b|\bwhich\b", smaller) and "if " not in smaller:
        return_type = "Name"
        cond_res = parse_single_condition(smaller)
        if cond_res:
            c_col, c_op, c_targ = cond_res
            return f'FILTER_NAMES("{c_col}", "{c_op}", "{c_targ}")', return_type
        else:
            return "Error: Incomplete condition for names.", return_type

    if " if " in smaller or "check if" in smaller or smaller.startswith("if ") or "has a mark" in smaller:
        if " and " in smaller or smaller.endswith(" and"): operation = "AND"
        elif " or " in smaller or smaller.endswith(" or"): operation = "OR"
        else: operation = "IF"
    elif "average" in smaller or "mean" in smaller: operation = "AVERAGE"
    elif "sum" in smaller or "total" in smaller or "calculate the sum" in smaller: operation = "SUM"   
    elif re.search(r"\b(max|maximum|top|highest)\b", smaller): operation = "MAX"
    elif re.search(r"\b(min|minimum|lowest)\b", smaller): operation = "MIN"
    elif "number of" in smaller or "count of" in smaller or "how many" in smaller or "show the" in smaller: operation = "COUNT"

    smaller = re.sub(r"\bis\b", "=", smaller)
    column_letter = column(smaller)

    if "pie" in smaller: operation = "PIE"
    elif "bar" in smaller or "chart" in smaller: operation = "BAR"
    elif "scatter" in smaller: operation = "SCATTER"

    if (operation == "MAX" or operation == "MIN") and column_letter == "":
        if "student" in smaller or "class" in smaller: column_letter = "D"
    if operation in ["BAR", "PIE", "SCATTER"] and column_letter == "": column_letter = "D"
    if operation == "COUNT" and column_letter == "": column_letter = "C"
            
    if operation in ["AND", "OR"]:
        if smaller.endswith(" or") or smaller.endswith(" and"):
            return "Error: Incomplete condition after connector.", return_type
            
        split_keyword = " and " if " and " in smaller else " or "
        clean_smaller = smaller.replace("check if ", "").replace("if ", "")
        parts = clean_smaller.split(split_keyword)
        
        if len(parts) < 2 or parts[0].strip() == "" or parts[1].strip() == "": 
            return "Error: Missing condition.", return_type
            
        cond1_res = parse_single_condition(parts[0])
        cond2_res = parse_single_condition(parts[1].strip())
        if not cond1_res or not cond2_res: return "Error: Invalid conditions.", return_type
        
        c1_col, c1_op, c1_targ = cond1_res
        c2_col, c2_op, c2_targ = cond2_res
        
        row1, row2 = str(start_row), str(start_row)
        actual_op = "AND" if split_keyword == " and " else "OR"
        return f'IF({actual_op}({c1_col}{row1}{c1_op}{c1_targ},{c2_col}{row2}{c2_op}{c2_targ}),"{true_output}","{false_output}")', return_type
        
    elif operation == "IF":
        cond_res = parse_single_condition(smaller)
        if not cond_res: return "Error: Incomplete condition.", return_type
        c_col, c_op, c_targ = cond_res
        print(f"DEBUG: Using row: {start_row}")
        return f'IF({c_col}{start_row}{c_op}{c_targ},"{true_output}","{false_output}")', return_type

    if operation != "" and column_letter != "":
        extra_filter = ""
        if operation == "COUNT" or (operation in ["BAR", "PIE"] and column_letter == "C"):
            has_female = bool(re.search(r"\b(female|females)\b", smaller))
            has_male = bool(re.search(r"\b(male|males)\b", smaller))
            has_both = "both" in smaller
            
            if (has_female and has_male) or has_both: 
                extra_filter = ", 'both'"
            elif has_female: 
                extra_filter = ", 'f'"
            elif has_male: 
                extra_filter = ", 'm'"
        range_str = f"{column_letter}{start_row}:{column_letter}{end_row}"
        return f"{operation}({range_str}{extra_filter})", return_type
    else: 
        return "Error: System could not parse.", return_type


def read_excel_cell(file_path, sheet_name, cell):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        else:
            sheet = wb.active
            
        value = sheet[cell].value
        if value is None:
            return 0
        try:
            return float(str(value).strip()) if '.' in str(value) else int(float(str(value).strip()))
        except:
            return value
    except Exception as e:
        return 0

def evaluate_cell(cell):
    return read_excel_cell(file_path, sheet_name, cell)

def validate_expression(expr, tokens):
    # فحص القوس الناقص
    if expr.count('(') != expr.count(')'):
        return "Error: Missing parenthesis"
    
    # فحص القسمة على صفر (البحث عن /0)
    if '/0' in expr.replace(' ', ''):
        return "Error: Division by zero"
    
    # فحص الخلايا: إذا وجدنا أي توكن يبدأ برقم (مثلاً 2B بدل B2)
    for t in tokens:
        if re.match(r'^\d+[A-Za-z]+', t):
            return f"Error: Invalid cell name '{t}' (Cell cannot start with a number)"
            
    return None # لا توجد أخطاء

def get_brackets_number(range_token):
    return range_token.count('(') , range_token.count(')')

def deal_with_range(tokens):
    left_range = []
    right_range = []
    all_ranges = []
    all_ranges_cells_name = []

    operation = None
    for op in ['+', '-', '*', '/']:
        if op in tokens:
            operation = op
            break

    if operation is None:
        return None, "Error: No arithmetic operation found."

    operation_index = tokens.index(operation)

    for i in range(0, operation_index):
        left_range.append(tokens[i])
        
    for i in range(operation_index + 1, len(tokens)):
        right_range.append(tokens[i])
        
    all_ranges.append(left_range)
    all_ranges.append(right_range)
    
    for my_range in all_ranges:
        total_left , total_right = get_brackets_number(my_range)
        if total_left != total_right:
            return None, "Error: Mismatched Brackets"
        else:
            colon_token_index = my_range.index(':')
            range_start = my_range[colon_token_index - 1]
            range_end = my_range[colon_token_index + 1]
            
            range_end_ref_num = re.findall(r'\d+', range_end)
            range_start_ref_num = re.findall(r'\d+', range_start)
            range_ref_name = re.findall(r'[A-Za-z]', range_end)
        
            range_cells_names = []
            for i in range(int(range_start_ref_num[0]), int(range_end_ref_num[0]) + 1):
                range_cells_names.append(str(range_ref_name[0]) + str(i)) 
                
            all_ranges_cells_name.append(range_cells_names)
            
    return all_ranges_cells_name, operation


def ranges_times(first_range, first_range_values, second_range, second_range_values):
    results = []
    for i in range(len(first_range)):
        res = first_range_values[i] * second_range_values[i]
        results.append(f'{first_range[i]} * {second_range[i]} = {res}')
    return "\n".join(results) 
def ranges_divide(first_range, first_range_values, second_range, second_range_values):
    results = []
    for i in range(len(first_range)):
        val = second_range_values[i]
        res = first_range_values[i] / val if val != 0 else 0
        results.append(f'{first_range[i]} / {second_range[i]} = {res}')
    return "\n".join(results)

def ranges_add(first_range, first_range_values, second_range, second_range_values):
    results = []
    for i in range(len(first_range)):
        res = first_range_values[i] + second_range_values[i]
        results.append(f'{first_range[i]} + {second_range[i]} = {res}')
    return "\n".join(results)

def ranges_subtrac(first_range, first_range_values, second_range, second_range_values):
    results = []
    for i in range(len(first_range)):
        res = first_range_values[i] - second_range_values[i]
        results.append(f'{first_range[i]} - {second_range[i]} = {res}')
    return "\n".join(results)

def get_cells_from_tokens(tokens):
    if len(tokens) > 2 and re.match(r'[A-Za-z]\d+', str(tokens[1])) and tokens[2] == ':' and re.match(r'[A-Za-z]\d+', str(tokens[-2])):
        range_end_ref_num = re.findall(r'\d+', str(tokens[-2]))
        range_start_ref_num = re.findall(r'\d+', str(tokens[1]))
        range_ref_name = re.findall(r'[A-Za-z]', str(tokens[1]))
        range_cells_names = []
        for i in range(int(range_start_ref_num[0]), int(range_end_ref_num[0]) + 1):
            range_cells_names.append(str(range_ref_name[0]) + str(i)) 
        return [read_excel_cell(file_path, sheet_name, cell) for cell in range_cells_names]
    else:
        range_values = []
        for tok in tokens:
            if re.match(r'[A-Za-z]\d+', str(tok)):
                val = read_excel_cell(file_path, sheet_name, str(tok))
                range_values.append(val)
        return range_values


def get_cells_from_range_str(range_str):
    cells = re.findall(r'[A-Za-z]\d+', range_str)
    if len(cells) >= 2:
        start_cell, end_cell = cells[0], cells[-1]
        col_letter = re.findall(r'[A-Za-z]', start_cell)[0]
        start_num = int(re.findall(r'\d+', start_cell)[0])
        end_num = int(re.findall(r'\d+', end_cell)[0])
        return [f"{col_letter}{i}" for i in range(start_num, end_num + 1)]
    return cells

def sum_values(tokens):
    if tokens[0] != '(' or tokens[-1] != ')': return -1
    if len(tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')': return 0
    vals = get_cells_from_tokens(tokens)
    return sum([float(x) for x in vals if isinstance(x, (int, float))])

def min_values(tokens):
    if tokens[0] != '(' or tokens[-1] != ')': return -1
    if len(tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')': return 0
    vals = get_cells_from_tokens(tokens)
    num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
    return min(num_vals) if num_vals else 0

def max_values(tokens):
    if tokens[0] != '(' or tokens[-1] != ')': return -1
    if len(tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')': return 0
    vals = get_cells_from_tokens(tokens)
    num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
    return max(num_vals) if num_vals else 0

def average_values(tokens):
    if tokens[0] != '(' or tokens[-1] != ')': return -1
    if len(tokens) == 2 and tokens[0] == '(' and tokens[-1] == ')': return 0
    vals = get_cells_from_tokens(tokens)
    num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
    return sum(num_vals) / len(num_vals) if num_vals else 0

def evaluate_condition_3(condition, truepart, falsepart):
    res = "" 
    if re.match(r'\d+', condition[0]):
        val = float(condition[0])
        operation = condition[1]
        target = float(condition[2])
        if operation == '>': res = truepart[0] if float(condition[0]) > float(condition[2]) else falsepart[0]
        elif operation == '<': res = truepart[0] if float(condition[0]) < float(condition[2]) else falsepart[0]
        elif operation == '=': 
            str_val = str(val).strip().rstrip('0').rstrip('.')
            str_target = str(float(condition[2])).strip().rstrip('0').rstrip('.')
            res = truepart[0] if str_val == str_target else falsepart[0]
        
            
    elif re.match(r'[A-Za-z]\d+', condition[0]):
        val = read_excel_cell(file_path, sheet_name, condition[0])
        operation = condition[1]
        if operation == '>': res = truepart[0] if val > float(condition[2]) else falsepart[0]
        elif operation == '<': res = truepart[0] if val < float(condition[2]) else falsepart[0]
        elif operation == '=': res = truepart[0] if val == float(condition[2]) else falsepart[0]
        elif operation == '>=': res = truepart[0] if val >= float(condition[2]) else falsepart[0]
        elif operation == '<=': res = truepart[0] if val <= float(condition[2]) else falsepart[0]
    if not res: res = falsepart[0]
    print(f"DEBUG: Comparison Result is: {res}")
    
    final_res = str(res).replace('"', '').replace("'", "")
    return res

def evaluate_condition_7(condition, truepart, falsepart):
    func_tokens = condition[0:6]
    res = sum_values(func_tokens)
    operation = condition[6]
    compare_value = float(condition[7])
    
   
    if operation == '>': res = truepart[0] if res > compare_value else falsepart[0]
    elif operation == '<': res = truepart[0] if res < compare_value else falsepart[0]
    elif operation == '=': res = truepart[0] if res == compare_value else falsepart[0]
    elif operation == '>=': res = truepart[0] if res >= compare_value else falsepart[0]
    elif operation == '<=': res = truepart[0] if res <= compare_value else falsepart[0]
    else: res = falsepart[0] 
    
   
    res = str(res).replace('"', '').replace("'", "")
    return res
    
def if_function(tokens):
    # 1. تنظيف قائمة التوكينز من 'IF' و '(' إن وجدوا في البداية
    while tokens and (tokens[0].upper() == 'IF' or tokens[0] == '('):
        tokens.pop(0)
    
    # 2. تنظيف الـ ')' من النهاية
    if tokens and tokens[-1] == ')':
        tokens.pop(-1)

    level = 0
    comma_indices = []
    for i, token in enumerate(tokens):
        if token == '(': level += 1
        elif token == ')': level -= 1
        elif token == ',' and level == 0: comma_indices.append(i)
    condition_tokens = tokens[:comma_indices[0]]
    truepart = "".join(tokens[comma_indices[0]+1 : comma_indices[1]]).replace("'", "").replace('"', "")
    falsepart = "".join(tokens[comma_indices[1]+1:]).replace("'", "").replace('"', "")
    if 'AND' in tokens or 'OR' in tokens:
        split_idx = tokens.index(')') + 1
        condition_tokens = tokens[:split_idx]
        remaining = tokens[split_idx + 1:] 
        comma_after_cond = remaining.index(',')
        truepart = [remaining[:comma_after_cond][0]]
        falsepart = [remaining[comma_after_cond+1:][0]]
        
        start_idx = tokens.index('(')
        end_idx = 0
        level = 0
        for i, t in enumerate(tokens):
            if t == '(': level += 1
            elif t == ')': level -= 1
            if level == 0 and i > start_idx:
                end_idx = i
                break
        
    
        full_condition = tokens[start_idx+1 : end_idx]
        
        
        expr_list = [str(read_excel_cell(file_path, sheet_name, t)) if re.match(r'[A-Za-z]+\d+', t) else t for t in full_condition]
        expr = "".join(expr_list).replace("AND", "all").replace("OR", "any").replace("^", "**")
        if "," in expr:
            expr = expr.replace(",", " and ")
        print(f"DEBUG: Perfect condition string: {expr}")
        res = eval(expr, {"__builtins__": None}, {
            "all": lambda *args: all(args), 
            "any": lambda *args: any(args),
        })    
        return truepart[0] if res else falsepart[0]
    level = 0
    comma_indices = []
    for i, token in enumerate(tokens):
        if token == '(': level += 1
        elif token == ')': level -= 1
        elif token == ',' and level == 0:
            comma_indices.append(i)

    
    
    condition = tokens[:comma_indices[0]] 
    if len(condition) == 2:
        cell, val = condition[0], condition[1]
        op = "<" if "<" in str(tokens) else ">" if ">" in str(tokens) else "="
        condition = [cell, op, val]
    truepart = tokens[comma_indices[0]+1 : comma_indices[1]]
    falsepart = tokens[comma_indices[1]+1:] 
    if len(condition) == 3: return evaluate_condition_3(condition, truepart, falsepart)
    elif len(condition) == 8: return evaluate_condition_7(condition, truepart, falsepart)

def and_function(tokens):
    tokens.pop(0)  # (
    tokens.pop(-1) # )
    comma_indices = [i for i, x in enumerate(tokens) if x == ","]
    condition1 = tokens[:comma_indices[0]]
    condition2 = tokens[comma_indices[0]+1:]
    
    res1 = False
    if re.match(r'\d+', condition1[0]):
        if condition1[1] == '>': res1 = float(condition1[0]) > float(condition1[2])
        if condition1[1] == '<': res1 = float(condition1[0]) < float(condition1[2])
        
    res2 = False
    if re.match(r'\d+', condition2[0]):
        if condition2[1] == '>': res2 = float(condition2[0]) > float(condition2[2])
        if condition2[1] == '<': res2 = float(condition2[0]) < float(condition2[2])
        
    final_res = True
    if res1 == False or res2 == False: final_res = False
    return final_res


def run_statistical_function(expression_str):
    fn_name = re.findall(r'[A-Za-z]+', expression_str)[0].upper()
    range_match = re.search(r'\((.*?)\)', expression_str)
    if not range_match: return "❌ Error: Invalid syntax in expression."
    
    parts = [p.strip().replace("'", "").replace('"', '') for p in range_match.group(1).split(',')]
    cells = get_cells_from_range_str(parts[0])
    filter_value = parts[1] if len(parts) > 1 else None
    vals = [read_excel_cell(file_path, sheet_name, c) for c in cells]
        
    if fn_name == 'SUM': res = sum([float(x) for x in vals if isinstance(x, (int, float))])
    elif fn_name == 'AVERAGE': 
        num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
        res = round(sum(num_vals) / len(num_vals), 2) if num_vals else 0.0
    elif fn_name == 'MIN': 
        num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
        res = min(num_vals) if num_vals else 0
    elif fn_name == 'MAX': 
        num_vals = [float(x) for x in vals if isinstance(x, (int, float))]
        res = max(num_vals) if num_vals else 0
    elif fn_name == 'COUNT': 
        if filter_value == 'both':
            males = sum(1 for x in vals if str(x).strip().lower() == 'm')
            females = sum(1 for x in vals if str(x).strip().lower() == 'f')
            return f"✅ Result: Total Students = {males + females} (Males: {males}, Females: {females})"
        elif filter_value: 
            res = sum(1 for x in vals if str(x).strip().lower() == filter_value.lower())
        else: 
            res = sum(1 for x in vals if x is not None and x != 0)
            
    return f"✅ Result: {res}"

def draw_chart(expression_str):
    fn_name = re.findall(r'[A-Za-z]+', expression_str)[0].upper()
    range_match = re.search(r'\((.*?)\)', expression_str)
    if not range_match: return None
    raw_content = range_match.group(1)
    parts = [p.strip() for p in raw_content.split(',')]
    
    if len(parts) >= 2 and any(char.isdigit() for char in parts[1]) and 'both' not in parts[1].lower():
        left_cells = get_cells_from_range_str(parts[0])
        right_cells = get_cells_from_range_str(parts[1])
        
        cat = [str(read_excel_cell(file_path, sheet_name, c)).strip() for c in left_cells]
        val = []
        for c in right_cells:
            cell_value = read_excel_cell(file_path, sheet_name, c)
            try: val.append(float(str(cell_value).strip()))
            except: val.append(0.0)
            
        title_text = f"{fn_name} Chart: {parts[0]} vs {parts[1]}"
        
    else:
        is_both_gender = 'both' in raw_content.lower()
        clean_range = parts[0]
        cells = get_cells_from_range_str(clean_range)
        col_letter = re.findall(r'[A-Za-z]', cells[0])[0].upper()
        if col_letter == 'C' or is_both_gender:
            all_vals = [str(read_excel_cell(file_path, sheet_name, c)).strip().lower() for c in cells if read_excel_cell(file_path, sheet_name, c) is not None]
            counts = Counter(all_vals)
            cat = ["Male (m)", "Female (f)"]
            val = [counts.get('m', 0), counts.get('f', 0)]
            title_text = "Number of Male and Female Students"
        elif col_letter in ['D', 'B', 'E', 'F']: 
            cat = []
            val = []
            for c in cells:
                row_num = re.findall(r'\d+', c)[0]
                student_name = str(read_excel_cell(file_path, sheet_name, f"A{row_num}")) 
                cell_value = read_excel_cell(file_path, sheet_name, c)
                if student_name and student_name != "0":
                    cat.append(student_name)
                    try: val.append(float(str(cell_value).strip()))
                    except: val.append(0.0)
            title_text = f"Values of Column {col_letter} per Student"
        else:
            all_vals = [str(read_excel_cell(file_path, sheet_name, c)).strip().lower() for c in cells if read_excel_cell(file_path, sheet_name, c) is not None]
            counts = Counter(all_vals)
            cat = list(counts.keys())
            val = list(counts.values())
            title_text = f"Distribution of Column {col_letter}"
            
    fig = plt.figure(figsize=(7, 4))
    if fn_name == 'BAR': 
        plt.bar(cat, val, color=['skyblue', 'lightpink'] if len(cat)==2 else 'skyblue')
        plt.ylabel('Values')
    elif fn_name == 'PIE': 
        plt.pie(val, labels=cat, autopct='%1.1f%%', startangle=140, colors=['skyblue', 'lightpink'] if len(cat)==2 else None)
    elif fn_name == 'SCATTER': plt.scatter(cat, val, color='red', s=100)
    elif fn_name == 'PLOT': plt.plot(cat, val, marker='o', color='green')
    if fn_name != 'PIE': plt.grid(True, linestyle='--', alpha=0.5)
    plt.title(title_text)
    plt.xticks(rotation=45 if len(cat) > 3 else 0)
    plt.tight_layout()
    return fig

def execute_filter_names(col_letter, operator, target_str):
    df_local = pd.read_excel(file_path)
    matched_names = []
    
    try: target_num = float(str(target_str).strip())
    except:
        try: target_num = float(re.sub(r'[^\d.]', '', str(target_str)))
        except: target_num = None

    op_clean = str(operator).strip()
    if op_clean not in [">", "<", ">=", "<="]: op_clean = "="
    
    df_local.columns = [str(c).strip().lower() for c in df_local.columns]
    name_col = 'name' if 'name' in df_local.columns else df_local.columns[0]
    mark_col = 'mark' if 'mark' in df_local.columns else (df_local.columns[3] if len(df_local.columns) > 3 else df_local.columns[-1])

    for index, row in df_local.iterrows():
        s_name = str(row[name_col]).strip()
        
        if col_letter.upper() == "B": current_val = row.get('age', row.iloc[1])
        elif col_letter.upper() == "C": current_val = row.get('gender', row.iloc[2])
        elif col_letter.upper() == "E": current_val = row.get('weight', row.iloc[4])
        elif col_letter.upper() == "F": current_val = row.get('height', row.iloc[5])
        else: current_val = row[mark_col]

        try: current_num = float(str(current_val).strip())
        except:
            try: current_num = float(re.sub(r'[^\d.]', '', str(current_val)))
            except: current_num = None

        is_match = False
        if current_num is not None and target_num is not None:
            if op_clean == "=":     is_match = (current_num == target_num)
            elif op_clean == ">":   is_match = (current_num > target_num)
            elif op_clean == "<":   is_match = (current_num < target_num)
            elif op_clean == ">=":  is_match = (current_num >= target_num)
            elif op_clean == "<=":  is_match = (current_num <= target_num)

        if is_match: matched_names.append(s_name)
            
  
    if matched_names: 
        return f"✅ Result: {', '.join(matched_names)}"
    else: 
        return "❌ Result: No students found matching this condition."

def syntax_analyzer(expression_str):
    tokens = tokenize(expression_str)
    print(f"DEBUG: Initial Tokens: {tokens}")
    if not tokens:
        return "Invalid expression: Empty input"

    equal_token = tokens.pop(0)
    if equal_token != '=':
        return 'Syntax Error : Excel Expressions must Starts with ='
    if tokens and tokens[0].upper() == 'IF':
        return if_function(tokens)

    clean_expr = expression_str.replace("=", "").strip()
    
    # 1. حالة الفلترة
    if clean_expr.startswith("FILTER_NAMES"):
        params = re.findall(r'"(.*?)"', clean_expr)
        if len(params) == 3:
            return execute_filter_names(params[0], params[1], params[2])
        else:
            return "❌ Error: Invalid parameters for name filtering."
        
    # 2. حالة الرسم البياني
    if any(chart_name in clean_expr.lower() for chart_name in ['bar', 'scatter', 'plot', 'pie']):
        return draw_chart(clean_expr)
        
    # 3. حالة الدوال الإحصائية
    if any(stat_name in clean_expr.lower() for stat_name in ['sum', 'average', 'min', 'max', 'count']) and (':' in clean_expr or ',' in clean_expr):
        if tokens and tokens[0].upper() in ['SUM', 'MIN', 'MAX', 'AVERAGE', 'IF', 'AND']:
            function_name_token = tokens.pop(0)
            function_name = function_name_token.upper()          
            if function_name == 'SUM':  res = sum_values(tokens)
            elif function_name == 'MIN': res = min_values(tokens)
            elif function_name == 'MAX': res = max_values(tokens)
            elif function_name == 'AVERAGE': res = average_values(tokens)
            elif function_name == 'IF': res = if_function(tokens)
            elif function_name == 'AND': res = and_function(tokens) 
            return f"✅ Result: {res}"
        else:
            return run_statistical_function(clean_expr)

    # 4. حالة العمليات الحسابية
    if ':' in clean_expr and any(op in clean_expr for op in ['+', '-', '*', '/']) and '(' not in clean_expr:
        cells_range, operation = deal_with_range(tokens)
        first_range = cells_range[0]
        second_range = cells_range[1]
        first_range_values = [read_excel_cell(file_path, sheet_name, cell) for cell in first_range]
        second_range_values = [read_excel_cell(file_path, sheet_name, cell) for cell in second_range]
            
        if operation == '*': return ranges_times(first_range, first_range_values, second_range, second_range_values)
        elif operation == '/': return ranges_divide(first_range, first_range_values, second_range, second_range_values)
        elif operation == '+': return ranges_add(first_range, first_range_values, second_range, second_range_values)
        elif operation == '-': return ranges_subtrac(first_range, first_range_values, second_range, second_range_values)

    # 5. حالة الحسابات العادية
    if "sqrt" in clean_expr.lower():
        clean_expr = re.sub(r"sqrt\s*\((.*?)\)", r"(\1)**0.5", clean_expr.lower())
        clean_expr = re.sub(r"sqrt\s+([a-zA-Z0-9_]+)", r"(\1)**0.5", clean_expr)

    expr = clean_expr.replace('^', '**')
    cells = sorted(list(set(re.findall(r'[A-Za-z]\d+', expr))), key=len, reverse=True)
    for cell in cells:
        val = read_excel_cell(file_path, sheet_name, cell)
        if isinstance(val, (int, float)):
            expr = expr.replace(cell, str(val))
        else:
            expr = expr.replace(cell, f"'{val}'")
    
    def IF(c, t, f): return t if c else f
    def AND(*args): return all(args)
    def OR(*args): return any(args)
    print(f"DEBUG: Final expression string before eval: {expr}")
    try:
        res = eval(expr, {"__builtins__": None}, {'IF': IF, 'AND': AND, 'OR': OR})
        if isinstance(res, float): res = round(res, 2)
        return str(res) # النتيجة النهائية
    except Exception as e: 
        return f"❌ Error: {e}"


def ask_system(user_sentence):
    result_text = "" 
    
 
    if user_sentence.strip().startswith("="):
        try:
            return syntax_analyzer(user_sentence) # هنا نستخدم return مباشرة
        except Exception as e:
            return f"Error: {str(e)}"

 
    math_expr_str = math_expression(user_sentence)
    if math_expr_str:
        try:
           
            return syntax_analyzer("=" + math_expr_str)
        except Exception as e:
            return f"Error: {str(e)}"

 
    res = natural(user_sentence)
    generated_formula = res[0] if isinstance(res, tuple) else res
        
    if "Error" in generated_formula: 
        return "❌ Sorry, I couldn't understand that."
        
    try:
        final_result = syntax_analyzer("=" + generated_formula)
        return final_result 
    except Exception as e:
        return f"Error: {str(e)}"