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

  # 4) 永久设置默认文档路径（目录或 .docx 路径），并可立即初始化
  python add.py --set-path "D:\notes" --init
  python add.py --set-path "D:\notes\我的词汇.docx"

  # 5) 恢复桌面默认
  python add.py --reset-path

  # 6) 查询当前路径与首次使用状态
  python add.py --where

  # 可选: --document 指定文档路径（仅本次生效，不修改默认配置）

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
    CONFIG_FILE,
    ensure_document, save_document,
    add_word, add_phrase, add_synonym,
    set_default_path, reset_default_path,
    get_default_doc_path, is_first_run,
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
    ap.add_argument('--document', help='目标 .docx 路径（仅本次生效；不修改默认配置）')
    ap.add_argument('--input', help='包含词条的 JSON 文件路径')
    ap.add_argument('--json', dest='json_str', help='词条 JSON 字符串（对象或数组）')
    ap.add_argument('--init', action='store_true', help='仅初始化文档，不写入')
    ap.add_argument('--set-path', dest='set_path',
                    help='永久设置默认文档路径（目录或 .docx 完整路径），可配合 --init')
    ap.add_argument('--reset-path', action='store_true',
                    help='清除自定义路径，恢复为桌面默认')
    ap.add_argument('--where', action='store_true',
                    help='显示当前默认路径、是否首次使用、配置文件位置')
    args = ap.parse_args()

    # 查询当前配置（不做任何写入）
    if args.where:
        doc_path = get_default_doc_path()
        print(f'config_file: {CONFIG_FILE}')
        print(f'first_run: {"true" if is_first_run() else "false"}')
        print(f'document_path: {doc_path}')
        print(f'document_exists: {"true" if os.path.exists(doc_path) else "false"}')
        return

    # 恢复桌面默认
    if args.reset_path:
        p = reset_default_path()
        print(f'[ok] 已恢复默认桌面路径: {p}')
        if args.init:
            ensure_document(None)
            print(f'[ok] 文档已就绪')
        return

    # 设置永久默认路径
    if args.set_path:
        p = set_default_path(args.set_path)
        print(f'[ok] 默认文档路径已保存: {p}')
        if args.init:
            ensure_document(p)
            print(f'[ok] 文档已就绪: {p}')
        return

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
