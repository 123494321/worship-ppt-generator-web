import re

# Standardize quotes and special brackets
def normalize_string(s):
    if not s:
        return ""
    return s.replace('’', "'").replace('`', "'").replace('（', '(').replace('）', ')')

def parse_routine_tokens(routine_str):
    """
    악보 루틴 문자열을 파싱하여 개별 토큰 및 반복 횟수 정보 [(token, count), ...]를 반환합니다.
    (예: 'Intro(8) - V1 - V2 - P - C - I(8) - Bx4 - C - Tag*2 - (기도)')
    """
    if not routine_str:
        return []
        
    norm = normalize_string(routine_str)
    # Split by hyphen or comma or whitespace around separators
    chunks = re.split(r'\s*-\s*|\s*,\s*', norm.strip())
    
    parsed_items = []
    
    for c in chunks:
        c = c.strip()
        if not c:
            continue
            
        # Match tokens like 'Intro(8)', 'Bx4', 'Tag*4', 'C(잔잔)', '(1)', '(기도)', '(목소리로) C'
        # Check sub-tokens
        sub_tokens = re.findall(r'[A-Za-z0-9가-힣\'*]+(?:\([^)]*\))?|\([^)]*\)[A-Za-z0-9가-힣\'*]*|\([^)]*\)', c)
        if not sub_tokens:
            sub_tokens = [c]
            
        for st in sub_tokens:
            st = st.strip()
            if not st:
                continue
                
            # Extract repeat count if format is Part x N or Part * N (e.g., 'Cx2', 'Bx4', 'Tag*2', 'P1x2', 'Tag x2')
            repeat = 1
            repeat_match = re.search(r'[*xX](\d+)\s*$', st)
            token_clean = st
            if repeat_match:
                try:
                    repeat = int(repeat_match.group(1))
                    token_clean = re.sub(r'\s*[*xX]\d+\s*$', '', st).strip()
                except:
                    repeat = 1
            else:
                # Also handle 'Tag(2)' or '(Tag*2 잔잔)'
                tag_num_match = re.match(r'^(?:Tag|tag|TAG)\s*\((\d+)\)$', st)
                if tag_num_match:
                    try:
                        repeat = int(tag_num_match.group(1))
                        token_clean = 'Tag'
                    except:
                        pass
                        
            parsed_items.append((token_clean, repeat))
            
    return parsed_items

def get_base_part_key(token):
    """
    토큰에서 핵심 파트 키를 추출합니다.
    (예: 'V3(느리게)' -> 'V3', 'C\'' -> 'C\'', '(1)' -> 'Interlude', 'Bx4' -> 'B', '(기도)' -> 'Prayer')
    """
    token_str = normalize_string(token).strip()
    
    # 1. 마디 수 간주 (예: '(1)', '(2)', '(4)', '(6)', '(8)', '(10)', '(12)', '(16)', '1', '4' 등)
    if re.match(r'^\(\s*\d+\s*\)$', token_str) or re.match(r'^\d+$', token_str):
        return 'Interlude'
        
    # 2. Interlude 계열 (예: 'I(4)', 'I(8)', 'ITL1(4)', 'ITL2(8)', 'I(8/V)', 'I', 'Inter', 'Interlude')
    if re.match(r'^(?:ITL\d*|I|Inter|Interlude)(?:\([^)]*\))?$', token_str, re.IGNORECASE):
        return 'Interlude'
        
    # 3. 기도 / 멘트 / 묵상
    if any(k in token_str for k in ['기도', '멘트', '묵상', 'Prayer', 'prayer']):
        return 'Prayer'
        
    # 4. 간주 / Solo / 기타 악기
    if any(k in token_str for k in ['간주', '쉬고', 'Solo', 'solo', 'Guitar', 'guitar', 'Drum', 'drum']):
        return 'Interlude'
        
    # 5. 전주 / 후주
    if 'Intro' in token_str or 'intro' in token_str:
        return 'Intro'
    if 'Outro' in token_str or 'outro' in token_str or token_str.startswith('Out') or token_str == 'O' or re.match(r'^O\(\d+\)$', token_str):
        return 'Outro'
        
    # 6. Key up 단독 지시어
    if 'key' in token_str.lower() and 'up' in token_str.lower() and not any(p in token_str for p in ['C', 'V', 'P', 'B']):
        return 'Interlude'

    # 7. 괄호 수식어 제거 후 분석 (예: '(목소리로) C' -> 'C', '(수아)V' -> 'V', 'C(잔잔)' -> 'C', 'P(break)' -> 'P')
    clean = re.sub(r'\(.*?\)', '', token_str).strip()
    
    if not clean:
        # 괄호를 지웠더니 내용이 없는 경우
        if '후렴' in token_str:
            return 'C'
        elif any(d in token_str for d in '0123456789'):
            return 'Interlude'
        else:
            # 특정 가사 괄호인 경우 (예: '(나 기쁨의 춤추리)') -> 사용자 정의 문장 반환
            inner = token_str.strip('()').strip()
            return f"CUSTOM:{inner}"
            
    # Part Key Regex
    # C' / C’ / C1 / C2 / V1 / V2 / V' / P1 / P2 / B1 / B2 / Tag / Tag1 / Tag2
    if re.match(r'^[Vv]\d*(?:\')?$', clean):
        return clean.upper() # V, V1, V2, V3, V', V1'
    elif re.match(r'^[Pp]\d*(?:\')?$', clean):
        return clean.upper() # P, P1, P2
    elif re.match(r'^[Cc]\d*(?:\')?$', clean):
        return clean.upper() # C, C1, C2, C3, C'
    elif re.match(r'^[Bb]\d*(?:\')?$', clean):
        return clean.upper() # B, B1, B2
    elif clean.upper().startswith('TAG') or clean.upper() == 'T' or re.match(r'^TAG\d*$', clean, re.IGNORECASE):
        return 'Tag'
        
    return clean

def get_major_part_key(token):
    """
    'V1', 'V2', 'V3', 'V'' -> 'V', 'C1', 'C2', 'C'' -> 'C', 'Tag1' -> 'Tag'
    대분류 파트 키를 반환합니다. (가사 Fallback용)
    """
    base = get_base_part_key(token)
    if base.startswith('CUSTOM:'):
        return base
    if base in ['Intro', 'Outro', 'Interlude', 'Prayer', 'Tag']:
        return base
    if base.startswith('V'):
        return 'V'
    elif base.startswith('P'):
        return 'P'
    elif base.startswith('C'):
        return 'C'
    elif base.startswith('B'):
        return 'B'
    return base
