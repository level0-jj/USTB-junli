import re
import json
from docx import Document

def full_to_half(text):
    result = []
    for ch in text:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            result.append(' ')
        else:
            result.append(ch)
    return ''.join(result)

doc = Document('军事理论题库.docx')
lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
raw_text = '\n'.join(lines)
raw_text = full_to_half(raw_text)

# 预处理：合并独立数字题号
raw_text = re.sub(r'\n(\d{1,3})\n', r'\n\1、', raw_text)

# 统一答案标记
raw_text = re.sub(r'答\s*案\s*[是为]*\s*[：:]', '答案：', raw_text)
raw_text = re.sub(r'参考\s*答案\s*[：:]', '答案：', raw_text)

# 切分题目
block_pattern = re.compile(r'(?:^|\n)\s*(\d+)\s*[、.．)]\s*')
splits = list(block_pattern.finditer(raw_text))
blocks = []
for i, m in enumerate(splits):
    start = m.start()
    next_start = splits[i+1].start() if i+1 < len(splits) else len(raw_text)
    block = raw_text[start:next_start].strip().lstrip('\n')
    blocks.append(block)

print(f'切分得到 {len(blocks)} 个题目块')

questions = []
skip_reasons = []

def is_judge_options(options):
    """判断选项列表是否属于判断题（只有正确/错误类文本）"""
    if len(options) != 2:
        return False
    judge_texts = {'正确', '错误', '对', '错', '√', '×', '是', '否'}
    return all(opt[1].strip() in judge_texts for opt in options)

for block in blocks:
    lines_in_block = block.splitlines()
    if not lines_in_block:
        skip_reasons.append('空块')
        continue

    stem = re.sub(r'^\s*\d+\s*[、.．)]\s*', '', lines_in_block[0]).strip()

    # ---------- 提取选项（字母范围 A-E）----------
    option_letters = []
    option_texts = []
    for line in lines_in_block:
        if re.match(r'^\s*\d+\s*[、.．)]', line):
            continue
        if re.match(r'^\s*答案', line):
            continue
        m = re.match(r'^\s*([A-E])\s*[.．、：:\s]*\s*(.*)', line)
        if m:
            option_letters.append(m.group(1))
            option_texts.append(m.group(2).strip().rstrip(';；'))

    options_raw = list(zip(option_letters, option_texts))

    # 如果选项数量不在 2~5 之间，用全文正则兜底
    if len(options_raw) < 2 or len(options_raw) > 5:
        opt_regex = re.findall(
            r'([A-E])\s*[.．、：:]\s*(.*?)(?=\s*[A-E]\s*[.．、：:]|答案|---|——|$)',
            block.replace('\n', ' ')
        )
        if opt_regex:
            options_raw = [(l, c.strip().rstrip(';；')) for l, c in opt_regex]

    # ---------- 提取答案 ----------
    answer_raw = None

    # 1) 括号内答案（字母或判断符号）
    bracket_match = re.search(r'[（(]\s*([A-E√×X╳对错正确错误](?:\s*[A-E√×X╳对错正确错误])*)\s*[）)]', block)
    if bracket_match:
        answer_raw = re.sub(r'\s+', '', bracket_match.group(1))
        stem = re.sub(r'[（(]\s*[A-E√×X╳对错正确错误]+(?:\s*[A-E√×X╳对错正确错误]+)*\s*[）)]', '（ ）', stem)

    # 2) “答案：”行
    line_match = re.search(r'答案\s*[：:]\s*([A-E对错正确错误√×X╳]+)', block)
    if not answer_raw and line_match:
        answer_raw = line_match.group(1)

    # 3) 破折号答案
    dash_match = re.search(r'[-—]+\s*([A-E]+)\s*', block)
    if not answer_raw and dash_match:
        answer_raw = dash_match.group(1)
        stem = re.sub(r'\s*[-—]+\s*[A-E]+\s*$', '', stem).strip()

    # 4) 判断题标记（题干尾）
    if not answer_raw:
        tail_match = re.search(r'[（(]?\s*([√×X╳对错正确错误])\s*[）)]?\s*$', stem)
        if tail_match:
            answer_raw = tail_match.group(1)
            stem = re.sub(r'\s*[（(]?\s*[√×X╳对错正确错误]\s*[）)]?\s*$', '', stem).strip()

    # 5) ():A 格式
    if not answer_raw:
        m_colon = re.search(r'[（(]\s*[）)]\s*[：:]\s*([A-E])', block)
        if m_colon:
            answer_raw = m_colon.group(1)
            stem = re.sub(r'[（(]\s*[）)]\s*[：:]\s*[A-E]', '（ ）', stem)

    # 6) 题干末尾孤立答案（仅当已有选项时）
    if not answer_raw and len(options_raw) >= 2:
        end_match = re.search(r'(?<=[^A-E])[A-E]{1,5}$', stem)
        if end_match:
            answer_raw = end_match.group(0)
            stem = re.sub(r'[A-E]{1,5}$', '', stem).rstrip()

    # ★★★ 新增 7) 题干括号内文本匹配选项 ★★★
    if not answer_raw and len(options_raw) >= 2:
        # 提取题干中所有圆括号/中文括号内的内容
        text_in_brackets = re.findall(r'[（(]([^）)]+)[）)]', stem)
        if text_in_brackets:
            # 对每个括号内容尝试匹配选项文本（忽略首尾空格和标点）
            for bracket_text in text_in_brackets:
                clean_bracket = bracket_text.strip(' ,，;；。、').strip()
                for letter, opt_text in options_raw:
                    clean_opt = opt_text.strip(' ,，;；。、').strip()
                    if clean_bracket == clean_opt:
                        answer_raw = letter
                        # 将题干中该括号内容替换为空的“（ ）”
                        stem = stem.replace(f'（{bracket_text}）', '（ ）').replace(f'({bracket_text})', '（ ）')
                        break
                if answer_raw:
                    break

    if not answer_raw:
        skip_reasons.append(f'无答案: {block[:60]}...')
        continue

    answer_clean = answer_raw.strip().upper()

    # ---------- 题型判定 ----------
    if len(options_raw) >= 2 and is_judge_options(options_raw):
        qtype = '判断'
        if answer_clean in ['√', '对', '正确', 'A']:
            answer = '正确'
        elif answer_clean in ['×', 'X', '╳', '错', '错误', 'B']:
            answer = '错误'
        else:
            if answer_clean == 'A':
                answer = '正确'
            elif answer_clean == 'B':
                answer = '错误'
            else:
                skip_reasons.append(f'判断答案异常: {answer_clean}')
                continue
        options = ['正确', '错误']

    elif len(options_raw) >= 2:
        if re.fullmatch(r'[A-E]+', answer_clean):
            if len(answer_clean) == 1:
                qtype = '选择'
                answer = answer_clean
            else:
                qtype = '多选'
                answer = answer_clean
            options = [f"{l}. {c}" for l, c in options_raw]
        else:
            skip_reasons.append(f'选项类题目答案非字母: {answer_clean}')
            continue

    else:
        if answer_clean in ['√', '对', '正确']:
            qtype = '判断'
            answer = '正确'
            options = ['正确', '错误']
        elif answer_clean in ['×', 'X', '╳', '错', '错误']:
            qtype = '判断'
            answer = '错误'
            options = ['正确', '错误']
        elif answer_clean in ['A', 'B']:
            qtype = '判断'
            answer = '正确' if answer_clean == 'A' else '错误'
            options = ['正确', '错误']
        elif re.fullmatch(r'[A-E]{2,}', answer_clean):
            skip_reasons.append(f'多选答案但无选项: {answer_clean}')
            continue
        else:
            skip_reasons.append(f'无选项且答案格式无法判定: {answer_clean}')
            continue

    stem = re.sub(r'\s+', ' ', stem).strip()
    questions.append({
        "id": len(questions) + 1,
        "type": qtype,
        "question": stem,
        "options": options,
        "answer": answer
    })

# ---------- 保存 ----------
with open('questions.json', 'w', encoding='utf-8') as f:
    json.dump(questions, f, ensure_ascii=False, indent=2)

print(f'单选: {sum(1 for q in questions if q["type"]=="选择")} 道')
print(f'多选: {sum(1 for q in questions if q["type"]=="多选")} 道')
print(f'判断: {sum(1 for q in questions if q["type"]=="判断")} 道')
print(f'\n丢弃题目数: {len(skip_reasons)}')
if skip_reasons:
    print('前10个丢弃原因:')
    for reason in skip_reasons[:10]:
        print(' -', reason)