# style_cache.py

from copy import copy

def get_cached_font(cell, font_cache): # 같은 원본 파일 안에서 동일한 폰트 서식을 재사용
    # 글꼴 이름뿐 아니라 크기·굵기·색상 등 전체 폰트 서식을 원본 서식 번호로 구분
    font_id = cell._style.fontId
    cached_font = font_cache.get(font_id)

    if cached_font is None:
        cached_font = copy(cell.font)
        font_cache[font_id] = cached_font

    return cached_font

def get_cached_border(cell, border_cache): # 부분 테두리를 포함한 모든 테두리 서식을 재사용
    # cell.has_style을 확인한 뒤 호출하며, 테두리 없음(borderId=0)도 캐싱
    # 선의 방향·굵기·색상 등 전체 테두리 설정을 원본 서식 번호로 구분
    border_id = cell._style.borderId
    cached_border = border_cache.get(border_id)

    if cached_border is None:
        cached_border = copy(cell.border)
        border_cache[border_id] = cached_border

    return cached_border

def get_cached_fill(cell, fill_cache): # "채우기 없음" 서식만 캐싱하고, 그 외 채우기는 기존처럼 복사
    fill = cell.fill

    # 색상·패턴·그라데이션 등 채우기가 적용된 경우에는 기존처럼 복사
    if fill.fill_type is not None:
        return copy(fill)

    # "채우기 없음"이어도 남아 있는 색상 속성이 다를 수 있으므로 원본 서식 번호로 구분
    fill_id = cell._style.fillId
    cached_fill = fill_cache.get(fill_id)

    if cached_fill is None:
        cached_fill = copy(fill)
        fill_cache[fill_id] = cached_fill

    return cached_fill

def get_cached_alignment(cell, alignment_cache): # 같은 원본 파일 안에서 동일한 정렬 서식을 재사용
    # 가로·세로 정렬, 줄 바꿈, 들여쓰기 등 전체 정렬 설정을 원본 서식 번호로 구분
    alignment_id = cell._style.alignmentId
    cached_alignment = alignment_cache.get(alignment_id)

    if cached_alignment is None:
        cached_alignment = copy(cell.alignment)
        alignment_cache[alignment_id] = cached_alignment

    return cached_alignment

def get_cached_protection(cell, protection_cache): # 같은 원본 파일 안에서 동일한 셀 보호 설정을 재사용
    # 잠금·수식 숨김 설정을 원본 서식 번호로 구분
    protection_id = cell._style.protectionId
    cached_protection = protection_cache.get(protection_id)

    if cached_protection is None:
        cached_protection = copy(cell.protection)
        protection_cache[protection_id] = cached_protection

    return cached_protection
