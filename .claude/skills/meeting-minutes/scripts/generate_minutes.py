#!/usr/bin/env python3
"""
generate_minutes.py
녹취록 기반 LGES 회의록 PPTX 생성 스크립트
"""

import argparse
import json
import os
import shutil
import zipfile
import re
from copy import deepcopy
from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def set_text_in_shape(shape_el, new_text):
    """텍스트박스의 첫 번째 run 텍스트를 교체. 나머지 run 제거."""
    txBody = shape_el.find('.//p:txBody', NS)
    if txBody is None:
        return
    paras = txBody.findall('a:p', NS)
    if not paras:
        return
    
    # 첫 번째 단락에만 텍스트 설정, 나머지 단락 제거
    first_para = paras[0]
    runs = first_para.findall('a:r', NS)
    
    if runs:
        # 첫 번째 run에 텍스트 설정
        t_el = runs[0].find('a:t', NS)
        if t_el is not None:
            t_el.text = new_text
        # 나머지 runs 제거
        for r in runs[1:]:
            first_para.remove(r)
    else:
        # run이 없으면 새로 생성
        r_el = etree.SubElement(first_para, '{http://schemas.openxmlformats.org/drawingml/2006/main}r')
        t_el = etree.SubElement(r_el, '{http://schemas.openxmlformats.org/drawingml/2006/main}t')
        t_el.text = new_text
    
    # 두 번째 이후 단락 제거
    for para in paras[1:]:
        txBody.remove(para)


def find_shape_by_text(tree, search_text):
    """특정 텍스트를 포함한 shape 요소 반환"""
    for sp in tree.findall('.//p:sp', NS):
        txBody = sp.find('.//p:txBody', NS)
        if txBody is None:
            continue
        full_text = ''.join(t.text or '' for t in txBody.findall('.//a:t', NS))
        if search_text in full_text:
            return sp
    return None


def find_all_shapes_with_text(tree):
    """모든 shape의 (id, name, text) 반환"""
    result = []
    for sp in tree.findall('.//p:sp', NS):
        nvSpPr = sp.find('p:nvSpPr/p:cNvPr', NS)
        sp_id = nvSpPr.get('id', '?') if nvSpPr is not None else '?'
        sp_name = nvSpPr.get('name', '') if nvSpPr is not None else ''
        txBody = sp.find('.//p:txBody', NS)
        if txBody is not None:
            text = ''.join(t.text or '' for t in txBody.findall('.//a:t', NS))
            result.append((sp_id, sp_name, text.strip()))
    return result


def build_attendee_string(attendees_raw):
    """
    attendees_raw: "LGES|FA기술담당|책임|강모원\nLGES|FA기술담당|팀장|최상훈"
    → "[LGES] FA기술담당 책임 강모원, 팀장 최상훈"
    회사별로 그룹화
    """
    lines = [l.strip() for l in attendees_raw.strip().split('\n') if l.strip()]
    company_groups = {}
    for line in lines:
        parts = line.split('|')
        if len(parts) < 4:
            continue
        company, org, title, name = parts[0], parts[1], parts[2], parts[3]
        key = f"{company}_{org}"
        if key not in company_groups:
            company_groups[key] = {'company': company, 'org': org, 'members': []}
        company_groups[key]['members'].append(f"{title} {name}")
    
    result_parts = []
    for key, group in company_groups.items():
        members_str = ', '.join(group['members'])
        result_parts.append(f"[{group['company']}] {group['org']} {members_str}")
    
    return '\n'.join(result_parts)


def update_slide(slide_xml_bytes, meta, attendees_str, items):
    """
    슬라이드 XML을 수정하여 반환
    
    meta: {date, weekday, time, location, dept, meeting_title}
    attendees_str: 포맷된 참석자 문자열
    items: [{"주관부서":str, "요약":str, "코멘트":str, "기한":str}, ...]
    """
    tree = etree.fromstring(slide_xml_bytes)
    
    # 디버그: 모든 shape 출력
    all_shapes = find_all_shapes_with_text(tree)
    
    # --- 1. 일시/장소 텍스트박스 ---
    date_str = f"{meta.get('date', 'YY.MM.DD')} ({meta.get('weekday', '요일')}) {meta.get('time', 'xx:00~xx:00')}, {meta.get('location', '장소')}"
    
    date_shape = find_shape_by_text(tree, '일시 / 장소')
    if date_shape:
        txBody = date_shape.find('.//p:txBody', NS)
        if txBody:
            for t in txBody.findall('.//a:t', NS):
                if '일시 / 장소' in (t.text or ''):
                    t.text = f"일시 / 장소 : {date_str}"
                    break
    
    # --- 2. 참석자 텍스트박스 ---
    attendee_shape = find_shape_by_text(tree, '참석자')
    if attendee_shape:
        txBody = attendee_shape.find('.//p:txBody', NS)
        if txBody:
            for t in txBody.findall('.//a:t', NS):
                if '참석자' in (t.text or '') and '직책' in (t.text or ''):
                    t.text = f"참석자 :  {attendees_str}"
                    break
    
    # --- 3. 날짜 레이블 (YY.MM.DD) ---
    date_label_shape = find_shape_by_text(tree, 'YY.MM.DD')
    if date_label_shape:
        txBody = date_label_shape.find('.//p:txBody', NS)
        if txBody:
            for t in txBody.findall('.//a:t', NS):
                if 'YY.MM.DD' in (t.text or ''):
                    t.text = meta.get('date_label', meta.get('date', 'YY.MM.DD'))
                    break
    
    # --- 4. 주관 부서 (상단) ---
    dept_shape = find_shape_by_text(tree, 'FA기술담당\n주관 부서')
    if not dept_shape:
        dept_shape = find_shape_by_text(tree, '주관 부서\n\n완료/보고 기한')
    # 주관 부서명 업데이트
    for sp in tree.findall('.//p:sp', NS):
        txBody = sp.find('.//p:txBody', NS)
        if txBody is None:
            continue
        text = ''.join(t.text or '' for t in txBody.findall('.//a:t', NS))
        if 'FA기술담당' in text and '주관 부서' in text:
            for t in txBody.findall('.//a:t', NS):
                if 'FA기술담당' in (t.text or ''):
                    t.text = meta.get('dept', 'FA기술담당')
                    break
    
    # --- 5. 회의록 제목 (헤더) ---
    title_shape = find_shape_by_text(tree, '[회의록]')
    if title_shape:
        txBody = title_shape.find('.//p:txBody', NS)
        if txBody:
            for t in txBody.findall('.//a:t', NS):
                if '[회의록]' in (t.text or ''):
                    t.text = f"[회의록] {meta.get('company', 'LGES')} {meta.get('dept', 'FA기술담당')} {meta.get('meeting_title', '업무보고')}"
                    break
    
    # --- 6. 이름 셀들 (6개 고정) ---
    name_shapes = []
    for sp in tree.findall('.//p:sp', NS):
        txBody = sp.find('.//p:txBody', NS)
        if txBody is None:
            continue
        text = ''.join(t.text or '' for t in txBody.findall('.//a:t', NS))
        if text.strip() == '이름':
            name_shapes.append(sp)
    
    attendee_lines = [l.strip() for l in attendees_str.split('\n') if l.strip()]
    # 이름만 추출
    all_names = []
    for line in attendees_raw_to_names(meta.get('_attendees_raw', '')):
        all_names.append(line)
    
    for i, sp in enumerate(name_shapes):
        name = all_names[i] if i < len(all_names) else ''
        txBody = sp.find('.//p:txBody', NS)
        if txBody:
            for t in txBody.findall('.//a:t', NS):
                if t.text and '이름' in t.text:
                    t.text = name
                    break
    
    # --- 7. 요약 항목 (최대 6개) ---
    # "주관 부서(cf. FA기술담당)"와 "요약 내용 N" 패턴으로 찾기
    summary_dept_shapes = []
    summary_content_shapes = []
    comment_shapes = []
    deadline_shapes = []
    
    for sp in tree.findall('.//p:sp', NS):
        txBody = sp.find('.//p:txBody', NS)
        if txBody is None:
            continue
        text = ''.join(t.text or '' for t in txBody.findall('.//a:t', NS)).strip()
        if re.match(r'^주관 부서\(cf\.', text):
            summary_dept_shapes.append(sp)
        elif re.match(r'^요약 내용 \d+$', text):
            summary_content_shapes.append(sp)
        elif text == 'Comment or':
            comment_shapes.append(sp)
        elif re.match(r'^YY\.MM$', text):
            deadline_shapes.append(sp)
    
    # 요약 항목 업데이트
    for i, item in enumerate(items[:6]):
        dept = item.get('주관부서', meta.get('dept', 'FA기술담당'))
        summary = item.get('요약', '')
        comment = item.get('코멘트', '')
        deadline = item.get('기한', 'YY.MM')
        
        if i < len(summary_dept_shapes):
            txBody = summary_dept_shapes[i].find('.//p:txBody', NS)
            if txBody:
                for t in txBody.findall('.//a:t', NS):
                    if t.text and '주관 부서' in t.text:
                        t.text = dept
                        break
        
        if i < len(summary_content_shapes):
            txBody = summary_content_shapes[i].find('.//p:txBody', NS)
            if txBody:
                for t in txBody.findall('.//a:t', NS):
                    if t.text and '요약 내용' in t.text:
                        t.text = summary
                        break
        
        if i < len(comment_shapes):
            txBody = comment_shapes[i].find('.//p:txBody', NS)
            if txBody:
                for t in txBody.findall('.//a:t', NS):
                    if t.text and 'Comment' in t.text:
                        t.text = comment
                        break
        
        if i < len(deadline_shapes):
            txBody = deadline_shapes[i].find('.//p:txBody', NS)
            if txBody:
                for t in txBody.findall('.//a:t', NS):
                    if t.text and 'YY.MM' in t.text:
                        t.text = deadline
                        break
    
    return etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)


def attendees_raw_to_names(raw):
    """참석자 raw 문자열에서 이름만 추출"""
    names = []
    lines = [l.strip() for l in raw.strip().split('\n') if l.strip()]
    for line in lines:
        parts = line.split('|')
        if len(parts) >= 4:
            names.append(parts[3])
    return names


def main():
    parser = argparse.ArgumentParser(description='LGES 회의록 PPTX 생성')
    parser.add_argument('--template', required=True, help='템플릿 PPTX 경로')
    parser.add_argument('--output', required=True, help='출력 PPTX 경로')
    parser.add_argument('--meta', required=True, help='JSON 형식 메타데이터')
    parser.add_argument('--attendees', required=True, help='참석자 raw 문자열 (파이프 구분)')
    parser.add_argument('--items', required=True, help='요약 항목 문자열 (인덱스|주관부서|요약|코멘트|기한 줄바꿈 구분)')
    args = parser.parse_args()
    
    meta = json.loads(args.meta)
    meta['_attendees_raw'] = args.attendees
    
    # 참석자 문자열 포맷
    attendees_str = build_attendee_string(args.attendees)
    
    # 요약 항목 파싱
    items = []
    for line in args.items.strip().split('\n'):
        if not line.strip():
            continue
        parts = line.split('|')
        if len(parts) >= 5:
            items.append({
                '주관부서': parts[1],
                '요약': parts[2],
                '코멘트': parts[3],
                '기한': parts[4],
            })
    
    # PPTX 복사 후 수정
    shutil.copy2(args.template, args.output)
    
    with zipfile.ZipFile(args.output, 'r') as zin:
        slide_xml = zin.read('ppt/slides/slide1.xml')
    
    updated_xml = update_slide(slide_xml, meta, attendees_str, items)
    
    # ZIP 내 슬라이드 교체
    import tempfile
    tmp_path = args.output + '.tmp'
    with zipfile.ZipFile(args.output, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'ppt/slides/slide1.xml':
                    zout.writestr(item, updated_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))
    
    os.replace(tmp_path, args.output)
    print(f"✅ 회의록 생성 완료: {args.output}")


if __name__ == '__main__':
    main()
