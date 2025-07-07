#!/usr/bin/env python3
"""
HWPX 파일의 특정 표 위치에 이미지 삽입하는 도구
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tempfile
import shutil
import sys
import base64
import uuid
from PIL import Image
import io

try:
    import lxml.etree as LXML_ET
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

class HWPXImageInserter:
    """HWPX 파일에 이미지 삽입하는 클래스"""
    
    def __init__(self):
        self.temp_dir = None
        self.image_counter = 0
        
    def insert_image_to_table(self, 
                            hwpx_file: str,
                            image_file: str, 
                            output_file: str,
                            table_index: int = 0,
                            row_index: int = 0,
                            col_index: int = 0,
                            image_options: dict = None):
        """
        HWPX 파일의 특정 표 셀에 이미지 삽입
        
        Args:
            hwpx_file: 원본 HWPX 파일 경로
            image_file: 삽입할 이미지 파일 경로
            output_file: 출력 HWPX 파일 경로
            table_index: 표 번호 (0부터 시작)
            row_index: 행 번호 (0부터 시작)
            col_index: 열 번호 (0부터 시작)
            image_options: 이미지 옵션 (크기, 정렬 등)
        
        Returns:
            dict: 처리 결과
        """
        
        print(f"🖼️  HWPX 이미지 삽입 시작")
        print(f"   원본: {hwpx_file}")
        print(f"   이미지: {image_file}")
        print(f"   출력: {output_file}")
        print(f"   위치: 표 {table_index}, 행 {row_index}, 열 {col_index}")
        
        # 기본 옵션 설정
        default_options = {
            'width': 100,  # mm 단위
            'height': 80,  # mm 단위
            'maintain_ratio': True,
            'alignment': 'center',  # left, center, right
            'vertical_alignment': 'middle'  # top, middle, bottom
        }
        if image_options:
            default_options.update(image_options)
        
        try:
            # 1. HWPX 파일 압축 해제
            self.temp_dir = tempfile.mkdtemp()
            hwpx_dir = Path(self.temp_dir) / "hwpx"
            
            with zipfile.ZipFile(hwpx_file, 'r') as zip_ref:
                zip_ref.extractall(hwpx_dir)
            
            print(f"✅ HWPX 압축 해제 완료")
            
            # 2. 이미지 파일 처리
            image_info = self._process_image(image_file, default_options)
            if not image_info['success']:
                return image_info
            
            # 3. 이미지를 HWPX에 추가
            self._add_image_to_hwpx(hwpx_dir, image_info)
            
            # 4. 표 찾기 및 이미지 삽입
            result = self._insert_image_to_table_cell(
                hwpx_dir, 
                table_index, 
                row_index, 
                col_index, 
                image_info,
                default_options
            )
            
            if not result['success']:
                return result
            
            # 5. 새로운 HWPX 파일 생성
            self._create_hwpx_file(hwpx_dir, output_file)
            
            print(f"✅ 이미지 삽입 완료: {output_file}")
            
            return {
                'success': True,
                'input_file': hwpx_file,
                'output_file': output_file,
                'image_file': image_file,
                'position': f"표 {table_index}, 행 {row_index}, 열 {col_index}",
                'image_size': f"{image_info['width']}x{image_info['height']}mm"
            }
            
        except Exception as e:
            print(f"❌ 이미지 삽입 실패: {e}")
            return {
                'success': False,
                'error': str(e)
            }
        finally:
            # 임시 디렉토리 정리
            if self.temp_dir and Path(self.temp_dir).exists():
                shutil.rmtree(self.temp_dir)
    
    def _process_image(self, image_file: str, options: dict) -> dict:
        """이미지 파일 처리 및 변환"""
        
        try:
            image_path = Path(image_file)
            if not image_path.exists():
                return {'success': False, 'error': f'이미지 파일을 찾을 수 없습니다: {image_file}'}
            
            # PIL로 이미지 열기
            with Image.open(image_path) as img:
                # RGB로 변환 (HWPX는 JPG 권장)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 크기 조정
                original_width, original_height = img.size
                
                if options['maintain_ratio']:
                    # 비율 유지하면서 크기 조정
                    ratio = min(options['width'] / (original_width * 0.264583), 
                              options['height'] / (original_height * 0.264583))  # px to mm 변환
                    new_width = int(original_width * ratio)
                    new_height = int(original_height * ratio)
                else:
                    # mm를 px로 변환 (72 DPI 기준)
                    new_width = int(options['width'] * 3.77953)  # mm to px
                    new_height = int(options['height'] * 3.77953)
                
                # 이미지 리사이즈
                img_resized = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                # JPG로 저장 (메모리에)
                img_buffer = io.BytesIO()
                img_resized.save(img_buffer, format='JPEG', quality=85)
                img_data = img_buffer.getvalue()
                
                # 고유 ID 생성
                image_id = f"image_{uuid.uuid4().hex[:8]}"
                
                return {
                    'success': True,
                    'id': image_id,
                    'data': img_data,
                    'width': options['width'],  # mm 단위
                    'height': options['height'],  # mm 단위
                    'pixel_width': new_width,
                    'pixel_height': new_height,
                    'format': 'jpg',
                    'size': len(img_data)
                }
            
        except Exception as e:
            return {'success': False, 'error': f'이미지 처리 실패: {e}'}
    
    def _add_image_to_hwpx(self, hwpx_dir: Path, image_info: dict):
        """이미지 파일을 HWPX 구조에 추가"""
        
        # BinData 폴더 생성
        bindata_dir = hwpx_dir / "BinData"
        bindata_dir.mkdir(exist_ok=True)
        
        # 이미지 파일 저장
        image_filename = f"{image_info['id']}.jpg"
        image_path = bindata_dir / image_filename
        
        with open(image_path, 'wb') as f:
            f.write(image_info['data'])
        
        print(f"📁 이미지 파일 추가: {image_filename}")
        
        # DocInfo/BinData.xml에 이미지 정보 추가
        self._update_bindata_xml(hwpx_dir, image_info, image_filename)
    
    def _update_bindata_xml(self, hwpx_dir: Path, image_info: dict, image_filename: str):
        """BinData.xml 파일 업데이트"""
        
        bindata_xml_path = hwpx_dir / "DocInfo" / "BinData.xml"
        
        # BinData.xml이 없으면 생성
        if not bindata_xml_path.exists():
            self._create_bindata_xml(bindata_xml_path)
        
        # XML 파싱
        if HAS_LXML:
            tree = LXML_ET.parse(str(bindata_xml_path))
            root = tree.getroot()
        else:
            tree = ET.parse(bindata_xml_path)
            root = tree.getroot()
        
        # BINDATASTORAGE 요소 찾기 또는 생성
        bindata_storage = root.find('BINDATASTORAGE')
        if bindata_storage is None:
            bindata_storage = ET.SubElement(root, 'BINDATASTORAGE')
        
        # BINDATA 요소 추가
        bindata = ET.SubElement(bindata_storage, 'BINDATA')
        bindata.set('id', image_info['id'])
        bindata.set('href', f"BinData/{image_filename}")
        bindata.set('type', 'jpg')
        bindata.set('size', str(image_info['size']))
        
        # XML 저장
        if HAS_LXML:
            tree.write(str(bindata_xml_path), encoding='utf-8', xml_declaration=True, pretty_print=True)
        else:
            self._indent_xml(root)
            tree.write(bindata_xml_path, encoding='utf-8', xml_declaration=True)
        
        print(f"📝 BinData.xml 업데이트 완료")
    
    def _create_bindata_xml(self, bindata_xml_path: Path):
        """기본 BinData.xml 파일 생성"""
        
        bindata_xml_path.parent.mkdir(exist_ok=True)
        
        root = ET.Element('BINDATASTORAGE')
        tree = ET.ElementTree(root)
        
        if HAS_LXML:
            tree.write(str(bindata_xml_path), encoding='utf-8', xml_declaration=True, pretty_print=True)
        else:
            tree.write(bindata_xml_path, encoding='utf-8', xml_declaration=True)
    
    def _insert_image_to_table_cell(self, hwpx_dir: Path, table_index: int, row_index: int, col_index: int, image_info: dict, options: dict):
        """표 셀에 이미지 삽입"""
        
        # BodyText 폴더에서 Section 파일들 찾기
        bodytext_dir = hwpx_dir / "BodyText"
        section_files = sorted(bodytext_dir.glob("Section*.xml"))
        
        table_found = False
        current_table_index = 0
        
        for section_file in section_files:
            result = self._process_section_file(
                section_file, 
                table_index, 
                row_index, 
                col_index, 
                current_table_index, 
                image_info, 
                options
            )
            
            if result['found']:
                table_found = True
                current_table_index = result['next_table_index']
                print(f"✅ {section_file.name}에서 이미지 삽입 완료")
                break
            
            current_table_index = result['next_table_index']
        
        if not table_found:
            return {
                'success': False,
                'error': f'표 {table_index}를 찾을 수 없습니다. 총 {current_table_index}개의 표가 있습니다.'
            }
        
        return {'success': True}
    
    def _process_section_file(self, section_file: Path, target_table: int, target_row: int, target_col: int, current_table_index: int, image_info: dict, options: dict):
        """Section 파일에서 표 처리"""
        
        try:
            # XML 파싱
            if HAS_LXML:
                tree = LXML_ET.parse(str(section_file))
                root = tree.getroot()
            else:
                tree = ET.parse(section_file)
                root = tree.getroot()
            
            # 표(TABLE) 요소들 찾기
            tables = root.findall('.//TABLE')
            
            for table in tables:
                if current_table_index == target_table:
                    # 목표 표 찾음
                    result = self._insert_image_to_table(table, target_row, target_col, image_info, options)
                    
                    if result['success']:
                        # 수정된 XML 저장
                        if HAS_LXML:
                            tree.write(str(section_file), encoding='utf-8', xml_declaration=True, pretty_print=True)
                        else:
                            self._indent_xml(root)
                            tree.write(section_file, encoding='utf-8', xml_declaration=True)
                        
                        return {
                            'found': True,
                            'next_table_index': current_table_index + 1
                        }
                    else:
                        return {
                            'found': False,
                            'next_table_index': current_table_index + 1,
                            'error': result['error']
                        }
                
                current_table_index += 1
            
            return {
                'found': False,
                'next_table_index': current_table_index
            }
            
        except Exception as e:
            print(f"❌ {section_file.name} 처리 실패: {e}")
            return {
                'found': False,
                'next_table_index': current_table_index,
                'error': str(e)
            }
    
    def _insert_image_to_table(self, table_element, target_row: int, target_col: int, image_info: dict, options: dict):
        """표 요소에 이미지 삽입"""
        
        try:
            # 행(TR) 요소들 찾기
            rows = table_element.findall('.//TR')
            
            if target_row >= len(rows):
                return {
                    'success': False,
                    'error': f'행 {target_row}가 존재하지 않습니다. 총 {len(rows)}개 행이 있습니다.'
                }
            
            target_row_element = rows[target_row]
            
            # 셀(TC) 요소들 찾기
            cells = target_row_element.findall('.//TC')
            
            if target_col >= len(cells):
                return {
                    'success': False,
                    'error': f'열 {target_col}가 존재하지 않습니다. 총 {len(cells)}개 열이 있습니다.'
                }
            
            target_cell = cells[target_col]
            
            # 셀 내용 지우기 (선택사항)
            # target_cell.clear()
            
            # 이미지 요소 생성
            image_element = self._create_image_element(image_info, options)
            
            # 셀에 이미지 추가
            target_cell.append(image_element)
            
            print(f"📍 이미지를 표의 행 {target_row}, 열 {target_col}에 삽입")
            
            return {'success': True}
            
        except Exception as e:
            return {
                'success': False,
                'error': f'이미지 삽입 실패: {e}'
            }
    
    def _create_image_element(self, image_info: dict, options: dict):
        """이미지 XML 요소 생성"""
        
        # PICTURE 요소 생성
        picture = ET.Element('PICTURE')
        picture.set('id', image_info['id'])
        picture.set('href', f"BinData/{image_info['id']}.jpg")
        
        # 크기 설정 (HWPX 단위는 1/100mm)
        width_hwp = int(image_info['width'] * 100)
        height_hwp = int(image_info['height'] * 100)
        
        picture.set('width', str(width_hwp))
        picture.set('height', str(height_hwp))
        
        # 정렬 설정
        if options['alignment'] == 'center':
            picture.set('textAlign', 'center')
        elif options['alignment'] == 'right':
            picture.set('textAlign', 'right')
        else:
            picture.set('textAlign', 'left')
        
        # REVERSE 요소 추가 (이미지 참조)
        reverse = ET.SubElement(picture, 'REVERSE')
        reverse.set('id', image_info['id'])
        
        return picture
    
    def _create_hwpx_file(self, hwpx_dir: Path, output_file: str):
        """수정된 디렉토리로부터 새로운 HWPX 파일 생성"""
        
        print(f"📦 HWPX 파일 생성 중...")
        
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in hwpx_dir.rglob('*'):
                if file_path.is_file():
                    relative_path = file_path.relative_to(hwpx_dir)
                    zipf.write(file_path, relative_path)
        
        print(f"✅ HWPX 파일 생성 완료")
    
    def _indent_xml(self, elem, level=0):
        """XML 요소에 들여쓰기 추가"""
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for child in elem:
                self._indent_xml(child, level + 1)
            if not child.tail or not child.tail.strip():
                child.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i
    
    def list_tables_in_hwpx(self, hwpx_file: str):
        """HWPX 파일의 표 구조 분석"""
        
        print(f"📋 HWPX 표 구조 분석: {hwpx_file}")
        
        try:
            temp_dir = tempfile.mkdtemp()
            hwpx_dir = Path(temp_dir) / "hwpx"
            
            with zipfile.ZipFile(hwpx_file, 'r') as zip_ref:
                zip_ref.extractall(hwpx_dir)
            
            bodytext_dir = hwpx_dir / "BodyText"
            section_files = sorted(bodytext_dir.glob("Section*.xml"))
            
            table_index = 0
            
            for section_file in section_files:
                if HAS_LXML:
                    tree = LXML_ET.parse(str(section_file))
                    root = tree.getroot()
                else:
                    tree = ET.parse(section_file)
                    root = tree.getroot()
                
                tables = root.findall('.//TABLE')
                
                for table in tables:
                    rows = table.findall('.//TR')
                    print(f"📊 표 {table_index}: {len(rows)}행")
                    
                    for row_idx, row in enumerate(rows):
                        cells = row.findall('.//TC')
                        print(f"   행 {row_idx}: {len(cells)}열")
                    
                    table_index += 1
            
            print(f"총 {table_index}개의 표가 발견되었습니다.")
            
        except Exception as e:
            print(f"❌ 표 분석 실패: {e}")
        finally:
            if Path(temp_dir).exists():
                shutil.rmtree(temp_dir)

def main():
    """CLI 인터페이스"""
    
    if len(sys.argv) < 6:
        print("""
HWPX 표에 이미지 삽입 도구

사용법:
    python hwpx_image_inserter.py input.hwpx image.jpg output.hwpx table_index row_index col_index
    python hwpx_image_inserter.py --list input.hwpx

예시:
    python hwpx_image_inserter.py document.hwpx photo.jpg new_document.hwpx 0 1 2
    (첫 번째 표의 2행 3열에 photo.jpg 삽입)
    
    python hwpx_image_inserter.py --list document.hwpx
    (문서의 표 구조 분석)

옵션:
    --list: 표 구조만 분석하고 종료
        """)
        sys.exit(1)
    
    # 표 구조 분석 모드
    if sys.argv[1] == '--list':
        inserter = HWPXImageInserter()
        inserter.list_tables_in_hwpx(sys.argv[2])
        return
    
    # 이미지 삽입 모드
    input_hwpx = sys.argv[1]
    image_file = sys.argv[2]
    output_hwpx = sys.argv[3]
    table_index = int(sys.argv[4])
    row_index = int(sys.argv[5])
    col_index = int(sys.argv[6])
    
    # 파일 존재 확인
    for file_path in [input_hwpx, image_file]:
        if not Path(file_path).exists():
            print(f"❌ 파일이 존재하지 않습니다: {file_path}")
            sys.exit(1)
    
    # 이미지 삽입 실행
    inserter = HWPXImageInserter()
    
    # 이미지 옵션 설정
    image_options = {
        'width': 80,  # mm
        'height': 60,  # mm
        'maintain_ratio': True,
        'alignment': 'center'
    }
    
    result = inserter.insert_image_to_table(
        hwpx_file=input_hwpx,
        image_file=image_file,
        output_file=output_hwpx,
        table_index=table_index,
        row_index=row_index,
        col_index=col_index,
        image_options=image_options
    )
    
    if result['success']:
        print(f"\n✅ 이미지 삽입 완료!")
        print(f"   입력: {result['input_file']}")
        print(f"   출력: {result['output_file']}")
        print(f"   위치: {result['position']}")
        print(f"   크기: {result['image_size']}")
    else:
        print(f"\n❌ 이미지 삽입 실패!")
        print(f"   오류: {result['error']}")
        sys.exit(1)

if __name__ == "__main__":
    main()
