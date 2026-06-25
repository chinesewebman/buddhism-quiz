#!/usr/bin/env python3
"""
build_quiz.py — 金刚经导读自测题 docx → HTML 单页答题系统

输入: /Volumes/sata11-130XXXX1371/PC/精舍/金刚经导读自测题/*.docx (29 个)
输出: /Users/apple/study/buddhism/output/{NN}-{主题}.html × 29
     + /Users/apple/study/buddhism/output/index.html
     + /Users/apple/study/buddhism/output/quizzes.zip

题型支持: 单选/多选/判断/填空/简答/论述
"""
import os, sys, re, json, html
from pathlib import Path

# 需要 python-docx
sys.path.insert(0, '/Users/apple/hermes-agent/venv/lib/python3.11/site-packages')
from docx import Document

# ═══════════════════════════════════════════════════════════════
# 路径
# ═══════════════════════════════════════════════════════════════
SRC_DIR = Path('/Volumes/sata11-130XXXX1371/PC/精舍/金刚经导读自测题')
OUT_DIR = Path('/Users/apple/study/buddhism/output')
DATA_DIR = Path('/Users/apple/study/buddhism/data')
OUT_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)


# ═══════════════════════════════════════════════════════════════
# 解析 docx
# ═══════════════════════════════════════════════════════════════

Q_TYPE_RE = re.compile(r'【(单选题|多选题|判断题|填空题|简答题|论述题)】')
ANSWER_RE = re.compile(r'^答案\s*[：:]\s*([^\n]*)')
PARSE_RE = re.compile(r'^解析\s*[：:]\s*([^\n]*)')
REF_ANSWER_RE = re.compile(r'^参考答案\s*[：:]\s*([^\n]*)')
OPT_RE = re.compile(r'^\s*([A-F])\s*[、.,。:：]?\s*([^\n]+)')
JUDGE_RE = re.compile(r'^\s*[\(（]([√✓]?)\s*[\)）]\s*(正确|错误)')
NUM_RE = re.compile(r'^\s*(\d+)\s*[.．、]\s*')


def parse_docx(path: Path) -> dict:
    """返回: {'number': '033', 'title': '庄严佛土', 'questions': [...]}"""
    doc = Document(path)
    paras = [p.text for p in doc.paragraphs]

    # 元数据 (可能不在段 0, 找前 3 段内第一个以 '所属题库' 开头的)
    meta_idx = None
    for k in range(min(3, len(paras))):
        if '所属题库' in paras[k]:
            meta_idx = k
            break
    if meta_idx is None:
        raise ValueError(f'{path.name}: 找不到元数据段')
    # 去掉前导符号 (如 '·' ' ' 等)
    meta_line = re.sub(r'^[^\S\n]*[·•\s]*', '', paras[meta_idx])
    m = re.match(r'所属题库[：:]\s*(?:初级)?金刚经导读(\d+)\s*(.+)', meta_line)
    if not m:
        raise ValueError(f'{path.name}: 元数据格式不对: {paras[meta_idx]}')
    number, title = m.group(1), m.group(2).strip()
    if title.endswith('之') or title.endswith('：'):
        title = title.rstrip('：')

    questions = []
    cur = None  # 当前题 dict
    in_parse = False
    parse_lines = []

    def commit_parse():
        """把累积的 parse_lines 拼到 cur['analysis']"""
        nonlocal parse_lines
        if cur is not None and parse_lines:
            cur['analysis'] = '\n'.join(parse_lines).strip()
            parse_lines = []
        else:
            parse_lines = []

    # 跳过元数据
    start_i = meta_idx + 1

    for i, raw in enumerate(paras):
        if i < start_i:
            continue
        line = raw.rstrip()
        s = line.strip()
        if not s:
            continue

        # 简答题参考答案延续: 累积到下一题/解析/答案
        # (放在最前: 优先收尾当前 ref, 再判断本段是新题/答案/解析)
        if cur is not None and cur.get('_collecting_ref'):
            if Q_TYPE_RE.search(s) or ANSWER_RE.match(s) or REF_ANSWER_RE.match(s) or PARSE_RE.match(s):
                # 下一题或新答案/解析段开始, 收尾
                cur['answer'] = '\n'.join(cur['_ref_lines']).strip()
                del cur['_collecting_ref']
                del cur['_ref_lines']
                # fall through (让本段继续按正常流程处理)
            else:
                if s:
                    cur['_ref_lines'].append(s)
                continue

        # 题目起始
        m = Q_TYPE_RE.search(s)
        if m:
            qtype_early = m.group(1)  # 提前取, sub-loop 要用
            # 检查段内是否含 \n 嵌入了选项
            if '\n' in s:
                # 拆分: 第一行为题干 (可能含 Q_TYPE), 后续行为选项/答案/解析
                lines_in = s.split('\n')
                # 第一行 = 题干
                stem_line = lines_in[0]
                # 后续行可能含 A. B. ... 答案:... 解析:...
                # 先把后续行按前缀分发
                cur_for_sub = {'opts': [], 'extra_answer': None, 'extra_parse': []}
                for sub in lines_in[1:]:
                    sub_s = sub.strip()
                    if not sub_s:
                        continue
                    opt_m = OPT_RE.match(sub)
                    # 检查: 当前题是选择题 OR (cur 存在且是选择题) — 兼容首题
                    is_sub_choice = (qtype_early in ('单选题', '多选题')) or (cur is not None and cur.get('is_choice'))
                    if opt_m and is_sub_choice:
                        cur_for_sub['opts'].append((opt_m.group(1), opt_m.group(2).strip()))
                        continue
                    ans_m2 = ANSWER_RE.match(sub_s)
                    if ans_m2:
                        cur_for_sub['extra_answer'] = ans_m2.group(1).strip()
                        continue
                    judge_m2 = JUDGE_RE.match(sub)
                    if judge_m2 and judge_m2.group(1):  # 有 √
                        cur_for_sub['extra_answer'] = judge_m2.group(2)
                        continue
                    parse_m2 = PARSE_RE.match(sub_s)
                    if parse_m2:
                        cur_for_sub['extra_parse'].append(parse_m2.group(1).strip())
                        cur_for_sub['had_parse'] = True
                        continue
                    # 解析延续
                    if cur_for_sub.get('had_parse'):
                        cur_for_sub['extra_parse'].append(sub_s)
                # 替换 s 为只有 stem_line
                s = stem_line
                # 把 sub 内容挂到一个待处理队列 (用一段伪段), 后面再处理
                pending_sub = cur_for_sub
            else:
                pending_sub = None

            commit_parse()
            if cur is not None:
                questions.append(cur)
            qtype = m.group(1)
            # 提取题号 (可能有, 也可能没有)
            num_m = NUM_RE.match(s)
            cur = {
                'type': qtype,
                'number': int(num_m.group(1)) if num_m else None,
                'stem': s,  # 包含题型标记的整段
                'options': [],   # [(letter, text), ...]
                'answer': None,  # 字符串
                'analysis': '',
                'is_choice': qtype in ('单选题', '多选题'),
            }
            # 处理 pending_sub
            if pending_sub:
                for ltr, txt in pending_sub['opts']:
                    cur['options'].append((ltr, txt))
                if pending_sub['extra_answer'] is not None:
                    cur['answer'] = pending_sub['extra_answer']
                if pending_sub['extra_parse']:
                    parse_lines.extend(pending_sub['extra_parse'])
                    in_parse = True
                else:
                    in_parse = False  # 新题, 重置解析状态
                pending_sub = None
            else:
                in_parse = False  # 新题, 重置解析状态
            # 新题: 初始化 _pending_A 候选为 'unset' (待定)
            cur['_pending_A'] = 'unset'
            continue

        # 题干延续段: 暂存为 A 候选, 后续无前缀段 → 追加为 B, C, ...
        if (cur is not None and cur['is_choice'] and not in_parse
            and not cur['options']
            and cur.get('_pending_A') not in ('set', None)
            and not ANSWER_RE.match(s) and not REF_ANSWER_RE.match(s) and not PARSE_RE.match(s)
            and not JUDGE_RE.match(line) and not Q_TYPE_RE.search(s)
            and not re.match(r'^\s*[\(（][√✓]?\s*[\)）]', s)
            and not re.match(r'^\s*\d+\s*[.．、]', s)
            and not re.match(r'^\s*[A-F]\s*[、.,。:：]', s)
            and not re.match(r'^\s*[A-F][^、.,。:：\s]', s)):
            pa = cur.get('_pending_A')
            if pa == 'unset':
                cur['_pending_A'] = s.strip()
            elif isinstance(pa, str) and pa != 'unset':
                # 已经有 A 字符串, 又来无前缀段, 转为 list 模式
                cur['_pending_A'] = [pa, s.strip()]
            elif isinstance(pa, list):
                pa.append(s.strip())
            cur['_pending_extra'] = cur['_pending_A'] if isinstance(cur['_pending_A'], list) else None
            continue

        # 选项
        if cur is not None and cur['is_choice'] and not in_parse:
            opt_m = OPT_RE.match(line)
            if opt_m:
                # 如果遇到 B/C/D 选项, 且 A 候选暂存着, 先 commit
                pa = cur.get('_pending_A')
                if opt_m.group(1) != 'A' and pa and pa != 'unset' and pa != 'set' and not cur['options']:
                    # pa 可能是 str 或 list
                    if isinstance(pa, str):
                        if '\n' in pa:
                            parts = pa.split('\n')
                            cur['options'].append(('A', parts[0].strip()))
                            for sub in parts[1:]:
                                sub_s = sub.strip()
                                if not sub_s:
                                    continue
                                sub_opt = OPT_RE.match(sub)
                                if sub_opt:
                                    cur['options'].append((sub_opt.group(1), sub_opt.group(2).strip()))
                        else:
                            cur['options'].append(('A', pa))
                    elif isinstance(pa, list):
                        for idx, txt in enumerate(pa):
                            letter = chr(ord('A') + idx)
                            cur['options'].append((letter, txt))
                cur['_pending_A'] = 'set'  # 已处理
                opt_text = opt_m.group(2)
                ans_set = False
                # 检查同段内 (opt_text 之后) 是否有 "答案:..." (无 \n 分隔)
                ans_inline = re.search(r'[。.]?\s*答案\s*[：:]\s*([^\n]+)', opt_text)
                if ans_inline:
                    clean_text = opt_text[:ans_inline.start()].rstrip('。. \t')
                    cur['options'].append((opt_m.group(1), clean_text))
                    if cur['answer'] is None:
                        cur['answer'] = ans_inline.group(1).strip()
                    ans_set = True
                else:
                    cur['options'].append((opt_m.group(1), opt_text.strip()))
                # 检查 \n 后续段: 可能是"答案"或"另一个选项"
                rest = s[opt_m.end():]
                if '\n' in rest:
                    # 继续处理后续 sub 行 (用递归风格)
                    sub_lines = rest.split('\n')
                    for sub in sub_lines:
                        sub_s = sub.strip()
                        if not sub_s:
                            continue
                        # 试匹配另一个选项
                        sub_opt = OPT_RE.match(sub)
                        if sub_opt:
                            cur['options'].append((sub_opt.group(1), sub_opt.group(2).strip()))
                            continue
                        # 试匹配答案
                        sub_ans = ANSWER_RE.match(sub_s)
                        if sub_ans and cur['answer'] is None:
                            cur['answer'] = sub_ans.group(1).strip()
                            ans_set = True
                            continue
                        # 试匹配解析
                        sub_parse = PARSE_RE.match(sub_s)
                        if sub_parse:
                            parse_lines.append(sub_parse.group(1).strip())
                            in_parse = True
                            continue
                        # 解析延续
                        if in_parse:
                            parse_lines.append(sub_s)
                continue

        # 答案 (注意: 解析中的"答案"也可能出现, 但我们按顺序解析)
        ans_m = ANSWER_RE.match(s)
        if ans_m and cur is not None and cur['answer'] is None:
            commit_parse()
            cur['answer'] = ans_m.group(1).strip()
            in_parse = False
            # 同一段可能紧跟 \n 解析, 处理它
            rest = s[ans_m.end():]
            rest = re.sub(r'^\s*\n\s*', '', rest)
            parse_inline = re.match(r'^解析\s*[：:]\s*([^\n]+)', rest)
            if parse_inline:
                parse_lines.append(parse_inline.group(1).strip())
                in_parse = True
            continue

        # 判断题答案: 从 "(√) 正确/错误" 行提取
        judge_m = JUDGE_RE.match(line)
        if judge_m and judge_m.group(1) and cur is not None and cur.get('type') == '判断题' and cur['answer'] is None:
            cur['answer'] = judge_m.group(2)
            in_parse = False
            continue

        # 简答题/论述题的参考答案
        ref_m = REF_ANSWER_RE.match(s)
        if ref_m and cur is not None:
            commit_parse()
            # 参考答案: 同段内剩余 + 后续非空段都并入
            rest = s[ref_m.end():].strip()
            ref_lines = []
            if rest:
                ref_lines.append(rest)
            cur['_collecting_ref'] = True
            cur['_ref_lines'] = ref_lines
            cur['is_open'] = True
            in_parse = False
            continue

        # 解析起始
        parse_m = PARSE_RE.match(s)
        if parse_m and cur is not None:
            # 如果正在收集参考答案, 先收尾 (上面已经处理过, 这里再保险)
            if cur.get('_collecting_ref'):
                cur['answer'] = '\n'.join(cur['_ref_lines']).strip()
                del cur['_collecting_ref']
                del cur['_ref_lines']
            commit_parse()
            parse_lines.append(parse_m.group(1).strip())
            in_parse = True
            continue

        # 解析延续: 跟在解析后, 且不是下一题/答案/选项
        if in_parse and cur is not None:
            # 检查: 如果这一行以 [A-F] 开头, 说明不是解析
            if not OPT_RE.match(line):
                parse_lines.append(s)

    # 收尾
    if cur is not None and cur.get('_collecting_ref'):
        cur['answer'] = '\n'.join(cur['_ref_lines']).strip()
        del cur['_collecting_ref']
        del cur['_ref_lines']
    commit_parse()
    if cur is not None:
        questions.append(cur)

    # 给没题号的题自动编号 (1, 2, 3...)
    auto_n = 1
    for q in questions:
        if q['number'] is None:
            q['number'] = auto_n
        auto_n = max(auto_n, q['number'] + 1)
    # 编号去重 (033 文件的题号从 3 开始, 前面已用 1, 2 占位)
    # 修正: 如果 number 序列不是 1..N, 重映射为连续
    nums = [q['number'] for q in questions]
    if nums != list(range(1, len(nums) + 1)):
        # 重映射
        for i, q in enumerate(questions, 1):
            q['number'] = i

    return {'number': number, 'title': title, 'questions': questions}


# ═══════════════════════════════════════════════════════════════
# HTML 渲染
# ═══════════════════════════════════════════════════════════════

# 古典东方哲学风格 CSS (米白宣纸 + 墨色 + 赭石点缀)
CSS = r"""
:root {
  --paper: #f7f1e6;        /* 宣纸米白 */
  --paper-2: #ede4d0;       /* 略深 */
  --ink: #2a2622;           /* 墨色 */
  --ink-soft: #5e564c;      /* 浅墨 */
  --ink-light: #8a8278;     /* 灰墨 */
  --ochre: #a35d2a;         /* 赭石 */
  --ochre-soft: #c89165;    /* 浅赭 */
  --cinnabar: #b34340;      /* 朱砂 (用于错误) */
  --jade: #5e7a4a;          /* 玉青 (用于正确) */
  --line: #c9bfa9;          /* 边框线 */
  --line-soft: #e3d9c2;     /* 软边框 */
}
* { box-sizing: border-box; }
html, body {
  margin: 0; padding: 0;
  background: var(--paper);
  color: var(--ink);
  font-family: "Noto Serif SC", "Songti SC", "STSong", "宋体", serif;
  font-size: 16px;
  line-height: 1.75;
  -webkit-font-smoothing: antialiased;
}
body {
  background:
    radial-gradient(ellipse at 10% 10%, rgba(167,93,42,0.04) 0%, transparent 40%),
    radial-gradient(ellipse at 90% 80%, rgba(94,122,74,0.04) 0%, transparent 50%),
    var(--paper);
  min-height: 100vh;
}
header.page {
  max-width: 880px;
  margin: 0 auto;
  padding: 56px 32px 32px;
  text-align: center;
  border-bottom: 1px solid var(--line);
  position: relative;
}
header.page::before, header.page::after {
  content: "";
  position: absolute;
  left: 50%; transform: translateX(-50%);
  width: 60px; height: 1px;
  background: var(--ink-light);
}
header.page::before { top: 32px; }
header.page::after  { bottom: 0; }
.kicker {
  font-size: 13px;
  letter-spacing: 0.4em;
  color: var(--ink-light);
  margin-bottom: 16px;
  text-transform: none;
  font-family: "Kaiti SC", "STKaiti", "KaiTi", serif;
}
h1.title {
  font-family: "Kaiti SC", "STKaiti", "KaiTi", serif;
  font-size: 38px;
  font-weight: 500;
  margin: 0 0 8px;
  letter-spacing: 0.1em;
}
.subtitle {
  font-size: 17px;
  color: var(--ink-soft);
  margin: 8px 0 24px;
  font-family: "Kaiti SC", "STKaiti", serif;
}
.meta {
  display: flex;
  justify-content: center;
  gap: 24px;
  font-size: 13px;
  color: var(--ink-light);
  letter-spacing: 0.1em;
}
.meta .sep { color: var(--line); }

main {
  max-width: 880px;
  margin: 0 auto;
  padding: 32px;
}
.nav-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-bottom: 24px;
  padding: 12px;
  background: var(--paper-2);
  border-radius: 4px;
  position: sticky;
  top: 0;
  z-index: 10;
  border: 1px solid var(--line-soft);
}
.nav-pills a {
  display: inline-block;
  min-width: 28px;
  padding: 4px 8px;
  text-align: center;
  font-size: 13px;
  color: var(--ink-soft);
  text-decoration: none;
  border: 1px solid transparent;
  border-radius: 3px;
  cursor: pointer;
  font-family: "Kaiti SC", "STKaiti", serif;
}
.nav-pills a:hover { background: var(--paper); border-color: var(--line); }
.nav-pills a.answered { background: var(--jade); color: var(--paper); }
.nav-pills a.wrong    { background: var(--cinnabar); color: var(--paper); }
.nav-pills a.current  { border-color: var(--ochre); color: var(--ochre); font-weight: 600; }

.q-card {
  background: #fdfaf2;
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 28px 32px;
  margin: 0 0 24px;
  position: relative;
  box-shadow: 0 1px 0 rgba(94,86,76,0.04);
}
.q-card.correct { border-color: var(--jade); border-width: 2px; }
.q-card.wrong   { border-color: var(--cinnabar); border-width: 2px; }
.q-head {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  margin-bottom: 16px;
  font-family: "Kaiti SC", "STKaiti", serif;
}
.q-no {
  font-size: 22px;
  color: var(--ochre);
  font-weight: 500;
  letter-spacing: 0.05em;
}
.q-no .badge {
  display: inline-block;
  font-size: 12px;
  padding: 2px 10px;
  margin-left: 8px;
  background: var(--paper-2);
  color: var(--ink-soft);
  border: 1px solid var(--line);
  border-radius: 12px;
  letter-spacing: 0.1em;
  font-weight: 400;
  vertical-align: middle;
}
.q-no .badge.judge   { color: var(--ochre); border-color: var(--ochre-soft); }
.q-no .badge.fill    { color: var(--jade); border-color: var(--jade); }
.q-no .badge.essay   { color: var(--ink-soft); }
.q-stem {
  font-size: 17px;
  line-height: 1.85;
  margin-bottom: 20px;
  color: var(--ink);
}
.options { margin: 16px 0; }
.opt {
  display: block;
  padding: 10px 16px;
  margin: 6px 0;
  background: var(--paper);
  border: 1px solid var(--line-soft);
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.15s;
  position: relative;
}
.opt:hover { background: var(--paper-2); border-color: var(--line); }
.opt input { margin-right: 12px; vertical-align: middle; }
.opt .ltr {
  display: inline-block;
  width: 22px;
  font-family: "Kaiti SC", "STKaiti", serif;
  color: var(--ochre);
  font-weight: 500;
}
.opt.is-correct { background: #e8f0dd; border-color: var(--jade); color: var(--ink); }
.opt.is-correct::after {
  content: "✓";
  position: absolute; right: 12px; top: 50%;
  transform: translateY(-50%);
  color: var(--jade); font-weight: 700;
}
.opt.is-wrong { background: #f6dcd9; border-color: var(--cinnabar); }
.opt.is-wrong::after {
  content: "✗";
  position: absolute; right: 12px; top: 50%;
  transform: translateY(-50%);
  color: var(--cinnabar); font-weight: 700;
}
.fill-input, .essay-input {
  width: 100%;
  padding: 10px 14px;
  font-size: 16px;
  font-family: inherit;
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: 4px;
  margin: 8px 0;
  color: var(--ink);
}
.fill-input:focus, .essay-input:focus { outline: none; border-color: var(--ochre); }
.essay-input { min-height: 140px; line-height: 1.7; resize: vertical; }

.actions {
  display: flex;
  gap: 12px;
  margin-top: 16px;
  align-items: center;
}
.btn {
  padding: 8px 20px;
  font-size: 14px;
  font-family: "Kaiti SC", "STKaiti", serif;
  background: var(--paper);
  color: var(--ink);
  border: 1px solid var(--ink-light);
  border-radius: 3px;
  cursor: pointer;
  letter-spacing: 0.15em;
  transition: all 0.15s;
}
.btn:hover { background: var(--paper-2); }
.btn-primary {
  background: var(--ochre);
  color: var(--paper);
  border-color: var(--ochre);
}
.btn-primary:hover { background: #8c4d20; }
.btn-primary:disabled { background: var(--line); border-color: var(--line); cursor: not-allowed; }
.btn-ghost { border-color: var(--line); color: var(--ink-soft); }

.feedback {
  margin-top: 18px;
  padding: 16px 20px;
  border-radius: 4px;
  display: none;
  font-size: 15px;
  line-height: 1.7;
}
.feedback.show { display: block; }
.feedback.right {
  background: #e8f0dd;
  border-left: 3px solid var(--jade);
  color: var(--ink);
}
.feedback.wrong {
  background: #f6dcd9;
  border-left: 3px solid var(--cinnabar);
  color: var(--ink);
}
.feedback .verdict {
  font-weight: 600;
  font-size: 16px;
  margin-bottom: 6px;
  font-family: "Kaiti SC", "STKaiti", serif;
  letter-spacing: 0.1em;
}
.feedback .verdict::before { content: "· "; color: var(--ink-light); }
.feedback .ref-answer {
  display: block;
  margin-top: 10px;
  padding: 10px 14px;
  background: rgba(255,255,255,0.5);
  border-radius: 3px;
  font-size: 15px;
}
.feedback .ref-answer .lbl {
  font-size: 12px;
  letter-spacing: 0.2em;
  color: var(--ink-light);
  display: block;
  margin-bottom: 4px;
  font-family: "Kaiti SC", "STKaiti", serif;
}
.feedback .analysis-text {
  display: block;
  margin-top: 10px;
  white-space: pre-wrap;
  line-height: 1.85;
}

footer {
  max-width: 880px;
  margin: 0 auto;
  padding: 32px;
  text-align: center;
  font-size: 12px;
  color: var(--ink-light);
  letter-spacing: 0.2em;
  border-top: 1px solid var(--line-soft);
  font-family: "Kaiti SC", "STKaiti", serif;
}

/* 印刷感收尾 */
@media (max-width: 600px) {
  .q-card { padding: 20px 18px; }
  h1.title { font-size: 28px; }
  main { padding: 16px; }
}
"""

JS = r"""
(function(){
  const KEY = (window.QSET || 'quiz') + '.progress';

  function load() {
    try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch(e) { return {}; }
  }
  function save(data) {
    try { localStorage.setItem(KEY, JSON.stringify(data)); } catch(e) {}
  }

  function setCurrent(anchor) {
    document.querySelectorAll('.nav-pills a').forEach(a => a.classList.remove('current'));
    const el = document.querySelector('.nav-pills a[href="#' + anchor + '"]');
    if (el) el.classList.add('current');
  }

  function grade(qid, type, answer) {
    const card = document.getElementById(qid);
    if (!card) return;
    const fb = card.querySelector('.feedback');
    const correctEl = card.querySelector('.q-no .verdict-mark');
    const userAnswer = collect(card, type);

    // 标记作答
    if (userAnswer !== null && userAnswer !== '' && !(Array.isArray(userAnswer) && userAnswer.length === 0)) {
      const navA = document.querySelector('.nav-pills a[href="#' + qid + '"]');
      if (navA) navA.classList.add('answered');
    }

    // 判分
    let isRight = null;
    if (type === '简答题' || type === '论述题') {
      // 开放题: 不判分, 只显示参考答案
      isRight = 'open';
    } else if (type === '判断题') {
      // 判断题: A=正确, B=错误; answer 可能是字母或文字, 归一化后再比较
      const normJudge = v => v === 'A' ? '正确' : v === 'B' ? '错误' : v;
      isRight = (normJudge(userAnswer) === normJudge(answer));
    } else if (type === '单选题') {
      isRight = (userAnswer === answer);
    } else if (type === '多选题') {
      const u = Array.isArray(userAnswer) ? userAnswer.slice().sort().join('') : '';
      const a = (answer || '').split('').sort().join('');
      isRight = (u === a);
    } else if (type === '填空题') {
      isRight = (String(userAnswer || '').trim() === String(answer || '').trim());
    }

    // 渲染反馈
    fb.classList.add('show');
    fb.classList.remove('right', 'wrong');
    if (isRight === 'open') {
      fb.classList.add('right');
      fb.innerHTML = '<div class="verdict">参考答案</div>' +
        '<span class="analysis-text">' + escHtml(answer) + '</span>' +
        renderAnalysis(card);
      card.classList.add('correct');
    } else if (isRight) {
      fb.classList.add('right');
      fb.innerHTML = '<div class="verdict">正确</div>' + renderAnalysis(card);
      card.classList.add('correct');
    } else {
      fb.classList.add('wrong');
      fb.innerHTML = '<div class="verdict">错误 · 正确答案: ' + escHtml(answer) + '</div>' + renderAnalysis(card);
      card.classList.add('wrong');
      const navA = document.querySelector('.nav-pills a[href="#' + qid + '"]');
      if (navA) {
        navA.classList.remove('answered');
        navA.classList.add('wrong');
      }
    }

    // 标记正确选项
    if (type === '单选题' || type === '多选题') {
      const correctSet = new Set((answer || '').split(''));
      card.querySelectorAll('.opt').forEach(opt => {
        const ltr = opt.getAttribute('data-ltr');
        if (correctSet.has(ltr)) opt.classList.add('is-correct');
      });
    } else if (type === '判断题') {
      // answer 可能是 '正确'/'错误' 或 'A'/'B', 归一化为字母再比较
      const ansLetter = answer === '正确' ? 'A' : answer === '错误' ? 'B' : answer;
      card.querySelectorAll('.opt').forEach(opt => {
        if (opt.getAttribute('data-ltr') === ansLetter) opt.classList.add('is-correct');
      });
    }

    // 禁用重复提交
    card.querySelector('.btn-submit').disabled = true;
    card.querySelector('.btn-submit').textContent = '已提交';

    // 保存进度
    const data = load();
    data[qid] = { type, answer, userAnswer, isRight, ts: Date.now() };
    save(data);

    // 更新统计
    updateSummary();
  }

  function collect(card, type) {
    if (type === '单选题' || type === '判断题') {
      const sel = card.querySelector('input[type=radio]:checked');
      return sel ? sel.value : null;
    }
    if (type === '多选题') {
      const sels = card.querySelectorAll('input[type=checkbox]:checked');
      return Array.from(sels).map(s => s.value);
    }
    if (type === '填空题') {
      const inp = card.querySelector('.fill-input');
      return inp ? inp.value : '';
    }
    if (type === '简答题' || type === '论述题') {
      const inp = card.querySelector('.essay-input');
      return inp ? inp.value : '';
    }
    return null;
  }

  function renderAnalysis(card) {
    const a = card.getAttribute('data-analysis');
    if (!a) return '';
    return '<div class="ref-answer"><span class="lbl">解析</span>' +
      '<span class="analysis-text">' + escHtml(a) + '</span></div>';
  }

  function escHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function updateSummary() {
    const data = load();
    const cards = document.querySelectorAll('.q-card');
    let total = 0, correct = 0, wrong = 0, open_ = 0;
    cards.forEach(c => {
      const qid = c.id;
      const type = c.getAttribute('data-type');
      const r = data[qid];
      if (!r) return;
      total += 1;
      if (r.isRight === 'open') open_ += 1;
      else if (r.isRight) correct += 1;
      else wrong += 1;
    });
    const sum = document.getElementById('summary');
    if (sum) {
      sum.innerHTML = '已作答 <b>' + total + '</b>/' + cards.length +
        ' · 正确 <b style="color:var(--jade)">' + correct + '</b>' +
        ' · 错误 <b style="color:var(--cinnabar)">' + wrong + '</b>' +
        ' · 开放题 <b>' + open_ + '</b>';
    }
  }

  function bind() {
    document.querySelectorAll('.btn-submit').forEach(btn => {
      btn.addEventListener('click', () => {
        const qid = btn.getAttribute('data-qid');
        const type = btn.getAttribute('data-type');
        const answer = btn.getAttribute('data-answer');
        grade(qid, type, answer);
      });
    });
    document.querySelectorAll('.btn-reset').forEach(btn => {
      btn.addEventListener('click', () => {
        const qid = btn.getAttribute('data-qid');
        const data = load();
        delete data[qid];
        save(data);
        location.reload();
      });
    });
    // 平滑滚动 + current 高亮
    document.querySelectorAll('.nav-pills a').forEach(a => {
      a.addEventListener('click', e => {
        setCurrent(a.getAttribute('href').slice(1));
      });
    });
    // 恢复历史进度
    const data = load();
    Object.keys(data).forEach(qid => {
      const r = data[qid];
      const card = document.getElementById(qid);
      if (!card) return;
      const navA = document.querySelector('.nav-pills a[href="#' + qid + '"]');
      if (r.isRight === true) {
        card.classList.add('correct');
        if (navA) navA.classList.add('answered');
      } else if (r.isRight === false) {
        card.classList.add('wrong');
        if (navA) { navA.classList.remove('answered'); navA.classList.add('wrong'); }
      } else if (r.isRight === 'open') {
        card.classList.add('correct');
        if (navA) navA.classList.add('answered');
      }
    });
    updateSummary();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bind);
  } else { bind(); }
})();
"""


def render_question(q: dict, qid: str) -> str:
    """渲染一题 HTML 卡片"""
    type_label = {
        '单选题': '单选', '多选题': '多选', '判断题': '判断',
        '填空题': '填空', '简答题': '简答', '论述题': '论述',
    }[q['type']]

    badge_class = {
        '判断题': 'judge', '填空题': 'fill', '简答题': 'essay', '论述题': 'essay',
    }.get(q['type'], '')

    # 题干 (简答/论述保留【题型】标记, 其他题型去掉)
    stem = q['stem']
    if q['type'] not in ('简答题', '论述题'):
        stem = re.sub(r'^【[^】]+】\s*', '', stem)  # 去【题型】(单选/多选/判断/填空)
    elif not re.match(r'^\s*\d*\s*【(简答题|论述题)】', stem):
        # 简答/论述题缺【】标记时, 兜底自动补 (即使 regex 误删, build 也不丢)
        stem = f'【{q["type"]}】{stem.lstrip()}'
    stem = re.sub(r'^\s*\d+\s*[.．、]\s*', '', stem)  # 去题号
    stem_html = html.escape(stem).replace('\n', '<br>')

    parts = [
        f'<article class="q-card" id="{qid}" data-type="{q["type"]}" data-analysis="{html.escape(q["analysis"])}">',
        '<div class="q-head">',
        f'<span class="q-no">{q["number"]}.<span class="badge {badge_class}">{type_label}</span></span>',
        '</div>',
        f'<div class="q-stem">{stem_html}</div>',
    ]

    # 选项区
    if q['is_choice']:
        input_type = 'radio' if q['type'] == '单选题' else 'checkbox'
        parts.append('<div class="options">')
        for letter, text in q['options']:
            parts.append(
                f'<label class="opt" data-ltr="{letter}">'
                f'<input type="{input_type}" name="{qid}" value="{letter}"> '
                f'<span class="ltr">{letter}.</span> {html.escape(text)}'
                f'</label>'
            )
        parts.append('</div>')
    elif q['type'] == '判断题':
        parts.append('<div class="options">')
        for letter, text in [('A', '正确'), ('B', '错误')]:
            parts.append(
                f'<label class="opt" data-ltr="{letter}">'
                f'<input type="radio" name="{qid}" value="{letter}"> '
                f'<span class="ltr">{letter}.</span> {html.escape(text)}'
                f'</label>'
            )
        parts.append('</div>')
    elif q['type'] == '填空题':
        parts.append(f'<input class="fill-input" name="{qid}" placeholder="请输入答案">')
    elif q['type'] in ('简答题', '论述题'):
        placeholder = '请简要作答…' if q['type'] == '简答题' else '请详细论述…'
        parts.append(f'<textarea class="essay-input" name="{qid}" placeholder="{placeholder}"></textarea>')

    # 操作 + 反馈
    answer_attr = q.get('answer') or ''
    answer_attr = html.escape(answer_attr).replace("'", "&#39;")
    parts.append(
        f'<div class="actions">'
        f'<button class="btn btn-primary btn-submit" data-qid="{qid}" data-type="{q["type"]}" data-answer="{answer_attr}">提交</button>'
        f'<button class="btn btn-ghost btn-reset" data-qid="{qid}">重做</button>'
        f'</div>'
        f'<div class="feedback"></div>'
        f'</article>'
    )
    return '\n'.join(parts)


def render_html(qset: dict) -> str:
    """生成完整 HTML"""
    questions = qset['questions']
    n = qset['number']
    title = qset['title']

    # 题号导航
    pills = '\n'.join(
        f'<a href="#q{n}-{q["number"]}" data-qid="q{n}-{q["number"]}">{q["number"]}</a>'
        for q in questions
    )

    # 题目卡片
    cards = '\n'.join(
        render_question(q, f'q{n}-{q["number"]}') for q in questions
    )

    type_count = {}
    for q in questions:
        type_count[q['type']] = type_count.get(q['type'], 0) + 1
    type_summary = ' · '.join(f'{k} {v}' for k, v in type_count.items())

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>金刚经导读 · 第{n}讲 · {title}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;500;700&family=Ma+Shan+Zheng&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<header class="page">
  <div class="kicker">金刚经导读 · 自测题</div>
  <h1 class="title">第 {n} 讲</h1>
  <div class="subtitle">{title}</div>
  <div class="meta">
    <span>共 <b>{len(questions)}</b> 题</span>
    <span class="sep">·</span>
    <span>{type_summary}</span>
    <span class="sep">·</span>
    <span id="summary">已作答 0/{len(questions)}</span>
  </div>
</header>

<main>
  <nav class="nav-pills">{pills}</nav>
  {cards}
</main>

<footer>
  <p>业精于勤 · 行成于思</p>
</footer>

<script>window.QSET = "quiz-{n}";</script>
<script>{JS}</script>
</body>
</html>"""


# ═══════════════════════════════════════════════════════════════
# 索引页
# ═══════════════════════════════════════════════════════════════

def render_index(quizzes: list) -> str:
    """生成索引页"""
    cards = []
    for q in quizzes:
        cards.append(
            f'<a class="quiz-card" href="{q["number"]}-{slug(q["title"])}.html">'
            f'<div class="quiz-no">第 {q["number"]} 讲</div>'
            f'<div class="quiz-title">{html.escape(q["title"])}</div>'
            f'<div class="quiz-meta">共 {len(q["questions"])} 题</div>'
            f'</a>'
        )
    cards_html = '\n'.join(cards)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>金刚经导读 · 自测题集</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;500;700&family=Ma+Shan+Zheng&display=swap" rel="stylesheet">
<style>
:root {{
  --paper: #f7f1e6; --paper-2: #ede4d0; --ink: #2a2622;
  --ink-soft: #5e564c; --ink-light: #8a8278;
  --ochre: #a35d2a; --ochre-soft: #c89165;
  --jade: #5e7a4a; --line: #c9bfa9; --line-soft: #e3d9c2;
}}
* {{ box-sizing: border-box; }}
body {{
  margin: 0; padding: 0;
  background: var(--paper);
  color: var(--ink);
  font-family: "Noto Serif SC", "Songti SC", serif;
  min-height: 100vh;
}}
header {{
  max-width: 1080px; margin: 0 auto;
  padding: 80px 32px 48px;
  text-align: center;
  border-bottom: 1px solid var(--line);
}}
.kicker {{ font-size: 13px; letter-spacing: 0.4em; color: var(--ink-light); margin-bottom: 16px; font-family: "Kaiti SC", "STKaiti", serif; }}
h1 {{ font-family: "Kaiti SC", "STKaiti", serif; font-size: 44px; margin: 0; letter-spacing: 0.1em; font-weight: 500; }}
.subtitle {{ font-size: 16px; color: var(--ink-soft); margin-top: 12px; font-family: "Kaiti SC", "STKaiti", serif; }}
main {{ max-width: 1080px; margin: 0 auto; padding: 48px 32px; }}
.grid {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 16px;
}}
.quiz-card {{
  display: block;
  padding: 24px 22px;
  background: #fdfaf2;
  border: 1px solid var(--line);
  border-radius: 6px;
  text-decoration: none;
  color: inherit;
  transition: all 0.2s;
  position: relative;
}}
.quiz-card:hover {{
  border-color: var(--ochre);
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(94,86,76,0.08);
}}
.quiz-card::before {{
  content: "";
  position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: var(--ochre);
  opacity: 0;
  transition: opacity 0.2s;
}}
.quiz-card:hover::before {{ opacity: 1; }}
.quiz-no {{
  font-family: "Kaiti SC", "STKaiti", serif;
  font-size: 13px;
  color: var(--ochre);
  letter-spacing: 0.2em;
  margin-bottom: 8px;
}}
.quiz-title {{
  font-size: 19px;
  font-weight: 500;
  margin-bottom: 8px;
  line-height: 1.4;
}}
.quiz-meta {{ font-size: 13px; color: var(--ink-light); letter-spacing: 0.1em; }}
footer {{
  text-align: center; padding: 48px 32px;
  color: var(--ink-light); font-size: 12px; letter-spacing: 0.2em;
  border-top: 1px solid var(--line-soft);
  font-family: "Kaiti SC", "STKaiti", serif;
}}
</style>
</head>
<body>
<header>
  <div class="kicker">金刚经导读</div>
  <h1>自测题集</h1>
  <div class="subtitle">三十三讲至六十一讲 · 共 {len(quizzes)} 卷</div>
</header>
<main>
  <div class="grid">
    {cards_html}
  </div>
</main>
<footer>业精于勤 · 行成于思</footer>
</body>
</html>"""


def slug(s: str) -> str:
    """标题 → URL slug (简单版)"""
    s = re.sub(r'[^\w\u4e00-\u9fff]+', '-', s)
    return s.strip('-')


# ═══════════════════════════════════════════════════════════════
# 主流程
# ═══════════════════════════════════════════════════════════════

def main():
    docx_files = sorted(SRC_DIR.glob('*.docx'))
    if not docx_files:
        print(f'❌ 没找到 docx 文件: {SRC_DIR}')
        sys.exit(1)
    print(f'找到 {len(docx_files)} 个 docx 文件')

    quizzes = []
    for f in docx_files:
        try:
            qset = parse_docx(f)
            print(f'  ✓ {f.name}: {qset["number"]} {qset["title"]} ({len(qset["questions"])} 题)')
            quizzes.append(qset)
        except Exception as e:
            print(f'  ❌ {f.name}: {e}')

    # 排序
    quizzes.sort(key=lambda q: int(q['number']))

    # 兜底: 修复"is_choice=true + options=空 + _pending_A 是 list"的情况
    # (例: 060.json q11 — 候选在 _pending_A 但没 B/C/D 行触发 commit)
    fixed_count = 0
    for qset in quizzes:
        for q in qset['questions']:
            if (q.get('is_choice')
                and not q.get('options')
                and isinstance(q.get('_pending_A'), list)):
                pa = q['_pending_A']
                q['options'] = [
                    [chr(ord('A') + idx), txt] for idx, txt in enumerate(pa)
                ]
                q['_pending_A'] = 'set'
                q.pop('_pending_extra', None)
                fixed_count += 1
                print(f'  🔧 Auto-fixed {qset["number"]} q{q["number"]} '
                      f'(promoted {len(pa)} _pending_A items to options)')
    if fixed_count:
        print(f'  → Total auto-fixed: {fixed_count}')

    # 输出 HTML
    for qset in quizzes:
        # 数据 JSON (调试用)
        with open(DATA_DIR / f'{qset["number"]}.json', 'w', encoding='utf-8') as f:
            json.dump(qset, f, ensure_ascii=False, indent=2)
        # HTML
        html_str = render_html(qset)
        out_path = OUT_DIR / f'{qset["number"]}-{slug(qset["title"])}.html'
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_str)
        print(f'  📄 {out_path.name}')

    # 索引页
    idx_html = render_index(quizzes)
    with open(OUT_DIR / 'index.html', 'w', encoding='utf-8') as f:
        f.write(idx_html)
    print(f'  📄 index.html')

    # CSV 数据库导出 (便于以后导入 SQLite / MySQL / Postgres)
    import csv as csv_mod
    csv_path = OUT_DIR / 'quizzes.csv'
    with open(csv_path, 'w', encoding='utf-8', newline='') as f:
        w = csv_mod.writer(f)
        # 字段定义
        w.writerow([
            'quiz_no',          # 题库编号: 033, 034, ...
            'quiz_title',       # 主题
            'q_no',             # 题号 (1-based)
            'q_type',           # 题型: 单选/多选/判断/填空/简答/论述
            'stem',             # 题干
            'options',          # 选项: "A. text || B. text || ..."
            'answer',           # 答案
            'analysis',         # 解析
            'is_open',          # 是否开放题 (1=是, 不自动判分)
            'source_file',      # 源 docx 文件名
        ])
        for qset in quizzes:
            for q in qset['questions']:
                opts_str = ' || '.join(f'{ltr}. {txt}' for ltr, txt in q.get('options', []))
                w.writerow([
                    qset['number'],
                    qset['title'],
                    q['number'],
                    q['type'],
                    q['stem'],
                    opts_str,
                    q.get('answer') or '',
                    q.get('analysis') or '',
                    '1' if q.get('is_open') or q['type'] in ('简答题', '论述题') else '0',
                    f.name,
                ])
    print(f'  📊 {csv_path.name} ({csv_path.stat().st_size//1024} KB)')

    # zip 包
    import zipfile
    zip_path = OUT_DIR / 'quizzes.zip'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(OUT_DIR.glob('*.html')):
            zf.write(f, f.name)
    print(f'  📦 {zip_path.name} ({zip_path.stat().st_size//1024} KB)')


if __name__ == '__main__':
    main()
