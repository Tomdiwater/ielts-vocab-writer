# -*- coding: utf-8 -*-
"""
ielts-vocab-writer 统一写入入口

用法:
  # 1) 仅初始化文档（桌面生成《雅思阅读词汇.docx》）
  python add.py --init

  # 2) 从 JSON 文件写入（可含单个对象或数组）
  python add.py --input entries.json

  # 3) 直接传 JSON 字符串
  python add.py --json "{\"type\":\"word\",...}"

  # 可选: --document 指定文档路径（默认桌面）

JSON 字段说明:
  word:    {type:"word", section:"学术", word:"excess",
            phonetic:"/ɪkˈses/", pos_meaning:"n. 过量",
            supplements:["搭配...","词族..."]}
  phrase:  {type:"phrase", phrase:"take off", meaning:"起飞；脱下",
            supplements:[...]}
  synonym: {type:"synonym", relation:"synonym"|"antonym",
            terms:["situated in","be located in"], meaning:"位于",
            supplements:[...]}
"""
import argparse
import json
import os
import sys

# 保证从任意工作目录运行都能 import common
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from common import (
    ensure_document, save_document,
    add_word, add_phrase, add_synonym,
)


def process_entry(doc, entry):
    t = entry.get('type', '').lower()
    if t == 'word':
        return add_word(doc, entry)
    if t == 'phrase':
        return add_phrase(doc, entry)
    if t in ('synonym', 'antonym'):
        if t == 'antonym':
            entry['relation'] = 'antonym'
        entry.setdefault('relation', 'synonym')
        return add_synonym(doc, entry)
    raise ValueError(f'未知 type: {t}（应为 word / phrase / synonym）')


def main():
    ap = argparse.ArgumentParser(description='雅思阅读词汇录入工具')
    ap.add_argument('--document', help='目标 .docx 路径（默认桌面《雅思阅读词汇.docx》）')
    ap.add_argument('--input', help='包含词条的 JSON 文件路径')
    ap.add_argument('--json', dest='json_str', help='词条 JSON 字符串（对象或数组）')
    ap.add_argument('--init', action='store_true', help='仅初始化文档，不写入')
    args = ap.parse_args()

    # 初始化/定位文档
    doc_path = ensure_document(args.document)
    if args.init:
        print(f'[ok] 文档就绪: {doc_path}')
        return

    # 读取词条数据
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    elif args.json_str:
        data = json.loads(args.json_str)
    else:
        ap.error('需要 --input 或 --json 或 --init')

    entries = data if isinstance(data, list) else [data]

    doc = Document(doc_path)
    results = []
    for i, entry in enumerate(entries, 1):
        try:
            msg = process_entry(doc, entry)
            results.append(f'  [{i}] {msg}')
        except Exception as e:
            results.append(f'  [{i}] 失败: {e}')

    saved = save_document(doc, doc_path)
    print('\n'.join(results))
    print(f'\n[done] 已写入 {len(entries)} 条，文档: {saved}')


if __name__ == '__main__':
    main()
