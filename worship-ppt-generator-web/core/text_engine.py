import re
from core.routine_parser import parse_routine_tokens, get_base_part_key, get_major_part_key, normalize_string

DEFAULT_CONTI_TEMPLATE = ""

def parse_user_conti(text):
    if not text or not text.strip():
        return {
            "date_str": "미입력",
            "title": "미입력",
            "prayer_person": "미입력",
            "songs": []
        }
        
    result = {
        "date_str": "미입력",
        "title": "미입력",
        "prayer_person": "미입력",
        "songs": []
    }
    
    lines = text.split('\n')
    current_song = None
    current_part = None
    part_buffer_lines = []
    
    def flush_part_buffer():
        nonlocal part_buffer_lines
        if current_song and current_part:
            raw_part_text = '\n'.join(part_buffer_lines).strip()
            if raw_part_text:
                chunks = re.split(r'\n\s*\n+', raw_part_text)
                slides = []
                for chk in chunks:
                    c = chk.strip()
                    if c:
                        slides.append(c)
                norm_part = normalize_string(current_part).upper()
                if norm_part not in current_song["parts"]:
                    current_song["parts"][norm_part] = []
                current_song["parts"][norm_part].extend(slides)
        part_buffer_lines = []

    for line in lines:
        stripped = line.strip()
        
        if stripped.startswith('# '):
            header = stripped[2:].strip()
            date_match = re.search(r'(\d{4}[.\-]\d{1,2}[.\-]\d{1,2})', header)
            if date_match:
                result["date_str"] = date_match.group(1).replace('-', '.')
                clean_title = header.replace(date_match.group(0), '').strip()
                if clean_title:
                    result["title"] = clean_title
            else:
                result["title"] = header
            continue
            
        if stripped.startswith('기도:') or stripped.startswith('대표기도:'):
            prayer_val = re.sub(r'^(기도|대표기도)\s*:\s*', '', stripped).strip()
            if prayer_val:
                result["prayer_person"] = prayer_val
            continue
            
        if stripped.startswith('##'):
            flush_part_buffer()
            raw_title = stripped.lstrip('#').strip()
            clean_title = re.sub(r'^\d+\s*[.\-)]\s*', '', raw_title).strip()
            current_song = {
                "title": clean_title,
                "routine_raw": "",
                "routine_items": [],
                "parts": {}
            }
            result["songs"].append(current_song)
            current_part = None
            continue
            
        if stripped.startswith('루틴:') or stripped.startswith('루틴 :'):
            if current_song:
                routine_val = re.sub(r'^루틴\s*:\s*', '', stripped).strip()
                current_song["routine_raw"] = routine_val
                current_song["routine_items"] = parse_routine_tokens(routine_val)
            continue
            
        part_match = re.match(r'^\[([A-Za-z0-9가-힣\'`’]+)\]$', stripped)
        if part_match:
            flush_part_buffer()
            current_part = normalize_string(part_match.group(1)).strip().upper()
            continue
            
        if current_song and current_part:
            part_buffer_lines.append(line)
            
    flush_part_buffer()
    return result

def generate_plan_text_from_conti(parsed_data, include_title_blank=False, include_end_blank=True):
    songs = parsed_data.get("songs", [])
    if not songs:
        return ""
        
    plan_blocks = []
    
    date_str = parsed_data.get("date_str", "미입력")
    title_str = parsed_data.get("title", "청년 모임")
    
    def add_slide(content):
        clean_content = content.strip()
        if clean_content == "[BLANK]" or not clean_content:
            # 연속 중복 암전 방지
            if plan_blocks and plan_blocks[-1] == "[BLANK]":
                return
            plan_blocks.append("[BLANK]")
        else:
            plan_blocks.append(clean_content)

    # 1. 표지
    add_slide(f"[{date_str}] {title_str}")
    # 2. 찬양 시작 전 암전
    add_slide("[BLANK]")
    
    for s_idx, song in enumerate(songs):
        song_title = song.get("title", f"곡 {s_idx+1}")
        items = song.get("routine_items", [])
        parts = song.get("parts", {})
        
        # 곡 제목 슬라이드
        add_slide(f"[ {song_title} ]")
        
        # 전주 암전
        if include_title_blank:
            add_slide("[BLANK]")
            
        for (token, repeat_count) in items:
            base_key = get_base_part_key(token)
            major_key = get_major_part_key(token)
            
            if base_key in ['Intro', 'Outro']:
                continue
            elif base_key in ['Interlude', 'Prayer']:
                add_slide("[BLANK]")
                continue
            elif base_key.startswith('CUSTOM:'):
                custom_lyric = base_key[7:].strip()
                for _ in range(repeat_count):
                    add_slide(custom_lyric)
                continue
                
            # 파트 가사 찾기
            part_slides = None
            token_upper = normalize_string(token).upper()
            base_upper = base_key.upper()
            major_upper = major_key.upper()
            
            if token_upper in parts and parts[token_upper]:
                part_slides = parts[token_upper]
            elif base_upper in parts and parts[base_upper]:
                part_slides = parts[base_upper]
            elif major_upper in parts and parts[major_upper]:
                part_slides = parts[major_upper]
            elif base_key == 'Tag':
                # 본문에 [Tag] 가사가 별도로 정의되지 않은 경우:
                # 방송실 운영 관례(구간 반복 시 기존 화면 유지)에 따라 중복 슬라이드를 만들지 않고 직전 화면을 고정 유지
                continue
                    
            if part_slides:
                for _ in range(repeat_count):
                    for s_text in part_slides:
                        add_slide(s_text)
            else:
                if re.search(r'\d+', token) and ('(' in token or ')' in token):
                    add_slide("[BLANK]")
                else:
                    for _ in range(repeat_count):
                        add_slide(f"[{token} 가사를 입력하세요]")
                
        # 곡 종료 시 암전 (필수 규칙에 따라 무조건 삽입)
        add_slide("[BLANK]")
            
    # 찬양 종료 암전
    add_slide("[BLANK]")
    
    # 대표기도 슬라이드
    prayer_person = parsed_data.get("prayer_person", "청년")
    add_slide(f"대표기도 : {prayer_person}")
    
    return '\n\n'.join(plan_blocks)

def parse_plan_text_to_slides(plan_text):
    if not plan_text or not plan_text.strip():
        return []
        
    blocks = re.split(r'\n\s*\n+', plan_text.strip())
    slides = []
    
    for idx, block in enumerate(blocks):
        b = block.strip()
        if not b:
            continue
            
        lines = b.split('\n')
        header_line = lines[0].strip()
        
        # 하위 호환성: 이전 포맷의 [숫자페이지] 헤더가 남아있는 경우에도 안전하게 파싱
        if re.match(r'^\[\d+페이지\]', header_line):
            content_lines = lines[1:]
        else:
            content_lines = lines
            
        content_text = '\n'.join([c.strip() for c in content_lines if c.strip()]).strip()
        
        if content_text == "[BLANK]" or not content_text:
            slides.append({
                "index": len(slides) + 1,
                "type": "blank",
                "text": "",
                "preview": "⬛ [암전]",
                "lines_count": 0
            })
        else:
            lines_list = [c.strip() for c in content_lines if c.strip()]
            first_line = lines_list[0] if lines_list else ""
            slide_type = "lyric"
            
            if idx == 0 and any(k in content_text for k in ['모임', '예배', '기도', '202']):
                slide_type = "title"
            elif first_line.startswith('[') and first_line.endswith(']'):
                slide_type = "song_title"
            elif first_line.startswith('대표기도') or first_line.startswith('기도'):
                slide_type = "prayer"
                
            slides.append({
                "index": len(slides) + 1,
                "type": slide_type,
                "text": content_text,
                "preview": first_line[:30] + ('...' if len(first_line) > 30 else ''),
                "lines_count": len(lines_list)
            })
            
    return slides

def get_upcoming_sunday(base_date=None):
    """
    기준 날짜(기본 오늘)로부터 가장 가까운 다가오는 일요일(주일) 날짜를 반환합니다.
    (오늘이 일요일이면 오늘 날짜 반환)
    """
    import datetime
    if base_date is None:
        base_date = datetime.date.today()
    days_ahead = (6 - base_date.weekday()) % 7
    return base_date + datetime.timedelta(days=days_ahead)

def get_default_conti_template(preset_name=None):
    """
    다가오는 주일 날짜와 프리셋 정보를 반영한 기본 콘티 템플릿을 생성합니다.
    (사용자 작성 편의를 극대화한 미니멀 스켈레톤 형식)
    """
    sunday = get_upcoming_sunday()
    sunday_str = sunday.strftime("%Y.%m.%d")
    
    group_name = "LOGOS 청년 모임"
    if preset_name and "대예배" in preset_name:
        group_name = "주일 대예배"
    elif preset_name and "중고등부" in preset_name:
        group_name = "중고등부 예배"
        
    template = f"""# {sunday_str} {group_name}
기도: ㅇㅇㅇ 청년

## 1. 
루틴: 

[V]

[P]

[C]

## 2. 
루틴: 

[V]

[P]

[C]
"""
    return template

def validate_conti_text(parsed_data):
    """
    콘티 구문 분석 결과를 검사하여 경고 및 상태 목록을 반환합니다.
    """
    warnings = []
    infos = []
    
    date_str = parsed_data.get("date_str", "미입력")
    if date_str == "미입력":
        warnings.append("표지 날짜를 적어주세요. (예: `# 2026.09.13 청년 모임`)")
    else:
        infos.append(f"날짜: {date_str}")
        
    prayer = parsed_data.get("prayer_person", "미입력")
    if prayer == "미입력":
        warnings.append("대표기도자를 적어주세요. (예: `기도: ㅇㅇㅇ 청년`)")
    else:
        infos.append(f"기도자: {prayer}")
        
    songs = parsed_data.get("songs", [])
    if not songs:
        warnings.append("찬양곡이 없습니다. (`## 1. 찬양제목` 형식으로 작성하세요)")
    else:
        infos.append(f"곡 수: {len(songs)}곡")
        for idx, s in enumerate(songs):
            title = s.get("title", "").strip()
            display_title = f"'{title}'" if title else f"곡 {idx+1}"
            items = s.get("routine_items", [])
            parts = s.get("parts", {})
            
            if not title:
                warnings.append(f"곡 {idx+1}의 제목을 적어주세요.")
            if not items:
                warnings.append(f"{display_title}의 루틴이 누락되었습니다. (예: `루틴: V - C`)")
            if not parts:
                warnings.append(f"{display_title}의 파트별 가사를 적어주세요.")

            else:
                for (token, _) in items:
                    base = get_base_part_key(token).upper()
                    if base not in ['INTRO', 'OUTRO', 'INTERLUDE', 'PRAYER'] and not base.startswith('CUSTOM:'):
                        if base not in parts and base not in ['TAG', 'TAG1']:
                            warnings.append(f"{display_title}: 루틴의 '[{token}]' 파트 가사가 본문에 없습니다.")
                            break
                            
                for p_name, p_slides in parts.items():
                    for sl in p_slides:
                        lines_in_sl = [l for l in sl.split('\n') if l.strip()]
                        if len(lines_in_sl) >= 3:
                            warnings.append(f"{display_title} [{p_name}]: 한 슬라이드에 가사가 {len(lines_in_sl)}줄입니다. 더블 엔터로 1~2줄 분할을 권장합니다.")
                            break
                            
    return {"warnings": warnings, "infos": infos}
