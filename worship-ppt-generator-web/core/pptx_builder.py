import io
import os
import copy
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from core.cloud_config import (
    DIR_PRESETS,
    DIR_PPTX,
    ensure_data_directories,
    format_standard_filename,
    parse_standard_filename
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ensure_data_directories()
PRESET_DIR = DIR_PRESETS

def get_preset_info_dict():
    """
    assets/preset_ppts 디렉토리를 스캔하여 하위 폴더 및 단일 PPTX 파일 정보를 파싱합니다.
    반환 딕셔너리 구조:
    {
        "프리셋 표시 이름": {
            "template_path": "...",
            "cover_path": "..." or None,
            "cover_name": "2026 청년부 예배 표지.png" or None,
            "is_folder": True/False
        }
    }
    """
    if not os.path.exists(PRESET_DIR):
        os.makedirs(PRESET_DIR, exist_ok=True)
    
    presets = {}
    entries = sorted(os.listdir(PRESET_DIR))
    
    # 1. 하위 폴더 프리셋 탐색 (권장 구조)
    for entry in entries:
        entry_path = os.path.join(PRESET_DIR, entry)
        if os.path.isdir(entry_path):
            sub_files = os.listdir(entry_path)
            pptx_list = [f for f in sub_files if f.endswith('.pptx') and not f.startswith('~$')]
            if pptx_list:
                template_path = os.path.join(entry_path, pptx_list[0])
                # 이미지 파일 탐색 (.png, .jpg, .jpeg) - 최신 수정일 순으로 정렬하여 최신 표지 우선 적용
                img_list = [f for f in sub_files if f.lower().endswith(('.png', '.jpg', '.jpeg')) and not f.startswith('~$')]
                if img_list:
                    img_list.sort(key=lambda f: os.path.getmtime(os.path.join(entry_path, f)), reverse=True)
                cover_path = os.path.join(entry_path, img_list[0]) if img_list else None
                cover_name = img_list[0] if img_list else None
                presets[entry] = {
                    "template_path": template_path,
                    "cover_path": cover_path,
                    "cover_name": cover_name,
                    "is_folder": True
                }
                
    # 2. 루트 단일 .pptx 파일 탐색 (하위 호환성)
    for entry in entries:
        entry_path = os.path.join(PRESET_DIR, entry)
        if os.path.isfile(entry_path) and entry.endswith('.pptx') and not entry.startswith('~$'):
            if entry not in presets:
                base_name = os.path.splitext(entry)[0]
                cover_path = None
                cover_name = None
                for ext in ['.png', '.jpg', '.jpeg']:
                    for suffix in ['', '_cover']:
                        c_cand = f"{base_name}{suffix}{ext}"
                        c_path = os.path.join(PRESET_DIR, c_cand)
                        if os.path.exists(c_path):
                            cover_path = c_path
                            cover_name = c_cand
                            break
                    if cover_path:
                        break
                presets[entry] = {
                    "template_path": entry_path,
                    "cover_path": cover_path,
                    "cover_name": cover_name,
                    "is_folder": False
                }
                
    return presets

def get_preset_ppts():
    """
    등록된 모든 표준 프리셋 이름 목록을 반환합니다.
    """
    info_dict = get_preset_info_dict()
    return sorted(list(info_dict.keys()))

def find_preset_cover_image(preset_name):
    """
    지정된 프리셋의 표지 이미지 파일 경로를 반환합니다.
    """
    info_dict = get_preset_info_dict()
    info = info_dict.get(preset_name)
    if info:
        return info.get("cover_path")
    return None

def extract_cover_image_from_template(template_source):
    """
    직접 업로드된 PPTX 또는 템플릿 파일의 1번 슬라이드에서
    모든 형식(Picture 도형, BlipFill 도형, 슬라이드 배경 이미지, 그룹 내부 이미지, OpenXML ImagePart)의
    표지 이미지를 100% 탐지하여 (io.BytesIO(blob), filename) 튜플로 반환합니다.
    표지 이미지가 없으면 None을 반환합니다.
    """
    try:
        if hasattr(template_source, 'seek'):
            template_source.seek(0)
        prs = Presentation(template_source)
        if len(prs.slides) == 0:
            return None
        slide0 = prs.slides[0]
        
        ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
        }
        
        candidates = []
        
        # 1. slide0의 XML 트리 전체에서 a:blip 탐색 (도형 채우기, 캔바 Freeform, 슬라이드 배경 등 전수 검사)
        blip_nodes = slide0._element.findall('.//a:blip', ns)
        r_embed_key = f"{{{ns['r']}}}embed"
        for blip in blip_nodes:
            r_id = blip.get(r_embed_key)
            if r_id and r_id in slide0.part.rels:
                part = slide0.part.rels[r_id].target_part
                blob = getattr(part, 'blob', None)
                if blob and len(blob) > 0:
                    ctype = getattr(part, 'content_type', '')
                    ext = 'jpg' if ('jpeg' in ctype or 'jpg' in ctype) else 'png'
                    candidates.append((blob, ext, len(blob)))
                    
        # 2. slide0.part.rels 전체에서 image 관계 전수 검사 (직접 Picture shape 등 방어)
        for rel_id, rel in slide0.part.rels.items():
            if 'image' in rel.reltype:
                part = rel.target_part
                blob = getattr(part, 'blob', None)
                if blob and len(blob) > 0:
                    ctype = getattr(part, 'content_type', '')
                    ext = 'jpg' if ('jpeg' in ctype or 'jpg' in ctype) else 'png'
                    if not any(c[0] == blob for c in candidates):
                        candidates.append((blob, ext, len(blob)))
                        
        if not candidates:
            return None
            
        # 2KB 미만의 투명 더미 스페이서(Canva Freeform 더미 등)를 배제하고 실제 표지 사진 선택
        real_images = [c for c in candidates if c[2] > 2048]
        chosen = max(real_images if real_images else candidates, key=lambda x: x[2])
        
        buf = io.BytesIO(chosen[0])
        filename = f"1번_슬라이드_추출표지.{chosen[1]}"
        buf.name = filename
        return buf, filename
    except Exception:
        return None



def clean_font_family(name):
    cleaned = name.replace(" (Canva 전용)", "").replace(" (윈도우 기본/호환 100%)", "").replace(" (프리텐다드)", "").strip()
    if "프리젠테이션" in cleaned:
        return "프리젠테이션 Bold"
    if cleaned.endswith(" Bold"):
        cleaned = cleaned[:-5].strip()
    return cleaned

def clean_proto_shape(grp_element):
    """
    그룹 셰이프 내부의 불필요한 투명 더미 이미지(Freeform/blip r:embed)를 제거하여
    파워포인트 오픈 시 깨짐이나 빨간색 x 표시를 원천 차단합니다.
    """
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    for sp in list(grp_element.findall('p:sp', ns)):
        cNvPr = sp.find('p:nvSpPr/p:cNvPr', ns)
        if cNvPr is not None and 'Freeform' in cNvPr.get('name', ''):
            grp_element.remove(sp)
    return grp_element

def extract_prototypes_from_prs(prs):
    """
    템플릿 프레젠테이션의 슬라이드를 순회하며 [제목], [1줄 가사], [2줄 가사]
    원본 그룹 셰이프 및 텍스트 박스 좌표/서식을 프로토타입으로 추출합니다.
    """
    proto_title = None
    proto_1line = None
    proto_2line = None

    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }

    for slide in prs.slides:
        if len(slide.shapes) == 0:
            continue
        shp = slide.shapes[0]
        text = ""
        for sp in shp._element.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
            txBody = sp.find('p:txBody', ns)
            if txBody is not None:
                p_nodes = txBody.findall('a:p', ns)
                text = '\n'.join([
                    ''.join([r.find('a:t', ns).text for r in p.findall('a:r', ns) if r.find('a:t', ns) is not None])
                    for p in p_nodes
                ])
                break
        lines = [l.strip() for l in text.split('\n') if l.strip()]

        if len(lines) == 1:
            if any(k in lines[0] for k in ['LOGOS', '청년', '모임', '기도', '[']):
                if proto_title is None:
                    proto_title = clean_proto_shape(copy.deepcopy(shp._element))
            else:
                if proto_1line is None:
                    proto_1line = clean_proto_shape(copy.deepcopy(shp._element))
        elif len(lines) >= 2:
            if proto_2line is None:
                proto_2line = clean_proto_shape(copy.deepcopy(shp._element))

        if (proto_title is not None) and (proto_1line is not None) and (proto_2line is not None):
            break

    return proto_title, proto_1line, proto_2line

def build_praise_pptx(slides_data, font_name=None, font_size_pt=None, template_source=None, cover_image=None):
    """
    사용자가 지정한 기준 베이스 템플릿(또는 표준 프리셋)의
    원본 좌표, 여백, 배율, 폰트 종류, 글자 크기, 폰트 임베딩을 100% 온전히 딥클론하여
    가사 줄 수에 맞게 슬라이드를 증감/생성하는 스마트 클론 PPTX 빌더입니다.
    """
    if template_source is not None:
        prs = Presentation(template_source)
    else:
        presets = get_preset_info_dict()
        if presets:
            first_key = sorted(list(presets.keys()))[0]
            prs = Presentation(presets[first_key]["template_path"])
        else:
            prs = Presentation()
            prs.slide_width = Inches(20.0)
            prs.slide_height = Inches(11.25)

    proto_title, proto_1line, proto_2line = extract_prototypes_from_prs(prs)

    # 템플릿 슬라이드 초기화 (내용만 비움)
    sldIdLst = prs.slides._sldIdLst
    for i in range(len(sldIdLst) - 1, -1, -1):
        prs.part.drop_rel(sldIdLst[i].rId)
        del sldIdLst[i]

    blank_layout = prs.slide_layouts[6]
    ns = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }

    for slide_idx, slide_info in enumerate(slides_data):
        slide = prs.slides.add_slide(blank_layout)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)

        # 1번째 슬라이드(표지)에 표지 이미지가 제공된 경우:
        # 비율과 관계없이 1번 슬라이드 전체에 여백 없이 꽉 채워 삽입
        if slide_idx == 0 and cover_image is not None and slide_info.get("type") == "title":
            if hasattr(cover_image, 'seek'):
                cover_image.seek(0)
            slide.shapes.add_picture(cover_image, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
            continue

        text_content = slide_info.get("text", "").strip()
        # 암전 슬라이드인 경우 빈 슬라이드로 유지
        if slide_info.get("type") == "blank" or not text_content:
            continue

        lines = [l.strip() for l in text_content.split('\n') if l.strip()]

        # 프로토타입 매칭
        if slide_info.get("type") == "title" or (len(lines) == 1 and any(k in text_content for k in ['LOGOS', '모임', '기도'])):
            proto = proto_title if proto_title is not None else (proto_1line if proto_1line is not None else proto_2line)
        elif len(lines) == 1:
            proto = proto_1line if proto_1line is not None else (proto_title if proto_title is not None else proto_2line)
        else:
            proto = proto_2line if proto_2line is not None else (proto_1line if proto_1line is not None else proto_title)

        if proto is not None:
            cloned_elem = copy.deepcopy(proto)
            slide.shapes._spTree.append(cloned_elem)

            for sp in cloned_elem.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
                cNvSpPr = sp.find('p:nvSpPr/p:cNvSpPr', ns)
                if cNvSpPr is not None and cNvSpPr.get('txBox') == 'true':
                    txBody = sp.find('p:txBody', ns)
                    if txBody is not None:
                        p_nodes = txBody.findall('a:p', ns)
                        for idx, line_str in enumerate(lines):
                            if idx < len(p_nodes):
                                p_curr = p_nodes[idx]
                            else:
                                p_curr = copy.deepcopy(p_nodes[-1])
                                txBody.append(p_curr)

                            runs = p_curr.findall('a:r', ns)
                            if runs:
                                t = runs[0].find('a:t', ns)
                                if t is not None:
                                    t.text = line_str
                                for extra in runs[1:]:
                                    p_curr.remove(extra)
                                rPr = runs[0].find('a:rPr', ns)
                                if rPr is not None:
                                    if font_size_pt is not None:
                                        xml_sz = str(int(font_size_pt / 0.75 * 100))
                                        rPr.set('sz', xml_sz)
                                    if font_name is not None:
                                        family = clean_font_family(font_name)
                                        for tag in ['latin', 'ea', 'cs', 'sym']:
                                            el = rPr.find(f'a:{tag}', ns)
                                            if el is not None:
                                                el.set('typeface', family)

                        # 여분의 문단 제거
                        if len(p_nodes) > len(lines):
                            for extra_p in p_nodes[len(lines):]:
                                txBody.remove(extra_p)
                    break
        else:
            # 폴백: 표준 텍스트 박스
            txBox = slide.shapes.add_textbox(Inches(0.0), Inches(0.0), prs.slide_width, prs.slide_height)
            tf = txBox.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            fallback_family = clean_font_family(font_name) if font_name else "Pretendard"
            fallback_pt = font_size_pt if font_size_pt is not None else 60.0
            for idx, line_str in enumerate(lines):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER
                run = p.add_run()
                run.text = line_str
                run.font.name = fallback_family
                run.font.size = Pt(fallback_pt)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

    return prs

def get_pptx_bytes(prs):
    """
    Presentation 객체를 바이트 스트림으로 반환합니다.
    """
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

def save_pptx_to_archive(prs, preset_name="청년부", title="찬양 가사", date_str="20260101"):
    """
    완성된 Presentation 객체를 데이터 보관함/2_찬양_PPT/ 에 표준 파일명으로 자동 저장하고
    (saved_filepath, filename, pptx_bytes) 튜플을 반환합니다.
    """
    os.makedirs(DIR_PPTX, exist_ok=True)
    filename = format_standard_filename(preset_name, "PPT", title, date_str, "pptx")
    save_path = os.path.join(DIR_PPTX, filename)
    prs.save(save_path)
    pptx_bytes = get_pptx_bytes(prs).getvalue()
    return save_path, filename, pptx_bytes
