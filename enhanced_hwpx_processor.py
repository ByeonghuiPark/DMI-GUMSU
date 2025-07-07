import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
import json
import tempfile
import shutil
import csv
from datetime import datetime
from dataclasses import dataclass
from typing import Dict, List, Optional, Union
import logging

# 추가 패키지들
try:
    import lxml.etree as LXML_ET
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

try:
    import chardet
    HAS_CHARDET = True
except ImportError:
    HAS_CHARDET = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import colorlog
    HAS_COLORLOG = True
except ImportError:
    HAS_COLORLOG = False

try:
    import regex as advanced_re
    HAS_REGEX = True
except ImportError:
    import re as advanced_re
    HAS_REGEX = False

@dataclass
class ProcessingStats:
    """처리 통계"""
    files_processed: int = 0
    files_successful: int = 0
    files_failed: int = 0
    total_replacements: int = 0
    processing_time: float = 0.0
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None

class EnhancedHWPXProcessor:
    """향상된 HWPX 텍스트 처리기"""
    
    def __init__(self, log_level: str = "INFO"):
        self.temp_dir = None
        self.sections = []
        self.document_info = {}
        self.stats = ProcessingStats()
        
        # 로거 설정
        self._setup_logger(log_level)
        
        # 기능 가용성 로그
        self._log_feature_availability()
    
    def _setup_logger(self, log_level: str):
        """컬러 로거 설정"""
        self.logger = logging.getLogger("HWPXProcessor")
        self.logger.setLevel(getattr(logging, log_level.upper()))
        
        if not self.logger.handlers:
            if HAS_COLORLOG:
                handler = colorlog.StreamHandler()
                handler.setFormatter(colorlog.ColoredFormatter(
                    '%(log_color)s%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S'
                ))
            else:
                handler = logging.StreamHandler()
                handler.setFormatter(logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S'
                ))
            
            self.logger.addHandler(handler)
    
    def _log_feature_availability(self):
        """사용 가능한 기능들 로그"""
        features = {
            "LXML (고성능 XML)": HAS_LXML,
            "CharDet (인코딩 감지)": HAS_CHARDET,
            "TQDM (진행률 표시)": HAS_TQDM,
            "Pandas (데이터 분석)": HAS_PANDAS,
            "ColorLog (컬러 로그)": HAS_COLORLOG,
            "Regex (고급 정규식)": HAS_REGEX
        }
        
        self.logger.info("=== 사용 가능한 기능 ===")
        for feature, available in features.items():
            status = "✓" if available else "✗"
            self.logger.info(f"{status} {feature}")
    
    def detect_file_encoding(self, file_path: Path) -> str:
        """파일 인코딩 자동 감지"""
        if not HAS_CHARDET:
            return 'utf-8'
        
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # 첫 10KB만 읽어서 감지
                result = chardet.detect(raw_data)
                encoding = result['encoding']
                confidence = result['confidence']
                
                self.logger.debug(f"인코딩 감지: {encoding} (신뢰도: {confidence:.2f})")
                
                # 신뢰도가 낮으면 기본값 사용
                if confidence < 0.7:
                    self.logger.warning(f"인코딩 감지 신뢰도 낮음. UTF-8 사용")
                    return 'utf-8'
                
                return encoding or 'utf-8'
        except Exception as e:
            self.logger.warning(f"인코딩 감지 실패: {e}. UTF-8 사용")
            return 'utf-8'
    
    def load_reference_data(self, data_source: Union[str, Dict, List]) -> Dict[str, str]:
        """향상된 참조 데이터 로드"""
        self.logger.info(f"참조 데이터 로드 시작: {data_source}")
        
        if isinstance(data_source, str):
            file_path = Path(data_source)
            
            if not file_path.exists():
                raise FileNotFoundError(f"참조 데이터 파일을 찾을 수 없습니다: {file_path}")
            
            if file_path.suffix.lower() == '.csv':
                return self._load_csv_reference_enhanced(file_path)
            elif file_path.suffix.lower() == '.json':
                return self._load_json_reference_enhanced(file_path)
            elif file_path.suffix.lower() in ['.xlsx', '.xls'] and HAS_PANDAS:
                return self._load_excel_reference(file_path)
            else:
                raise ValueError(f"지원하지 않는 파일 형식: {file_path.suffix}")
        
        elif isinstance(data_source, dict):
            self.logger.info(f"딕셔너리에서 {len(data_source)}개 항목 로드")
            return data_source
        
        elif isinstance(data_source, list):
            result = dict(data_source)
            self.logger.info(f"리스트에서 {len(result)}개 항목 로드")
            return result
        
        else:
            raise ValueError("지원하지 않는 데이터 형식입니다.")
    
    def _load_csv_reference_enhanced(self, csv_path: Path) -> Dict[str, str]:
        """향상된 CSV 참조 데이터 로드"""
        encoding = self.detect_file_encoding(csv_path)
        reference_dict = {}
        
        try:
            if HAS_PANDAS:
                # Pandas를 사용한 로드
                df = pd.read_csv(csv_path, encoding=encoding)
                
                # 첫 두 컬럼을 검색어-치환어로 사용
                if len(df.columns) >= 2:
                    search_col = df.columns[0]
                    replace_col = df.columns[1]
                    
                    for _, row in df.iterrows():
                        search_term = str(row[search_col]).strip()
                        replace_term = str(row[replace_col]).strip()
                        
                        if search_term and search_term != 'nan':
                            reference_dict[search_term] = replace_term
                
                self.logger.info(f"Pandas로 CSV에서 {len(reference_dict)}개 항목 로드")
            else:
                # 기본 CSV 모듈 사용
                with open(csv_path, 'r', encoding=encoding) as f:
                    reader = csv.reader(f)
                    
                    # 헤더 처리
                    first_row = next(reader, None)
                    if first_row and not (first_row[0].lower() in ['search', '검색어', 'original']):
                        if len(first_row) >= 2:
                            reference_dict[first_row[0].strip()] = first_row[1].strip()
                    
                    # 데이터 행 처리
                    for row in reader:
                        if len(row) >= 2 and row[0].strip():
                            reference_dict[row[0].strip()] = row[1].strip()
                
                self.logger.info(f"CSV에서 {len(reference_dict)}개 항목 로드")
        
        except Exception as e:
            self.logger.error(f"CSV 로드 실패: {e}")
            raise
        
        return reference_dict
    
    def _load_excel_reference(self, excel_path: Path) -> Dict[str, str]:
        """엑셀 파일에서 참조 데이터 로드 (Pandas 필요)"""
        if not HAS_PANDAS:
            raise ImportError("엑셀 파일 처리를 위해 pandas가 필요합니다: pip install pandas openpyxl")
        
        try:
            df = pd.read_excel(excel_path)
            reference_dict = {}
            
            if len(df.columns) >= 2:
                search_col = df.columns[0]
                replace_col = df.columns[1]
                
                for _, row in df.iterrows():
                    search_term = str(row[search_col]).strip()
                    replace_term = str(row[replace_col]).strip()
                    
                    if search_term and search_term != 'nan':
                        reference_dict[search_term] = replace_term
            
            self.logger.info(f"엑셀에서 {len(reference_dict)}개 항목 로드")
            return reference_dict
        
        except Exception as e:
            self.logger.error(f"엑셀 로드 실패: {e}")
            raise
    
    def _load_json_reference_enhanced(self, json_path: Path) -> Dict[str, str]:
        """향상된 JSON 참조 데이터 로드"""
        encoding = self.detect_file_encoding(json_path)
        
        try:
            with open(json_path, 'r', encoding=encoding) as f:
                data = json.load(f)
            
            if isinstance(data, dict):
                self.logger.info(f"JSON에서 {len(data)}개 항목 로드")
                return data
            else:
                raise ValueError("JSON 파일은 딕셔너리 형태여야 합니다.")
        
        except Exception as e:
            self.logger.error(f"JSON 로드 실패: {e}")
            raise
    
    def search_and_replace_text(self, 
                              hwpx_file: str, 
                              reference_data: Union[str, Dict, List],
                              output_file: Optional[str] = None,
                              replacement_options: Optional[Dict] = None) -> Dict:
        """향상된 텍스트 검색 및 치환"""
        
        start_time = datetime.now()
        self.logger.info(f"파일 처리 시작: {hwpx_file}")
        
        # 기본 옵션 설정
        options = {
            'case_sensitive': False,
            'whole_word_only': False,
            'use_regex': False,
            'use_advanced_regex': HAS_REGEX,
            'max_replacements_per_term': -1,
            'preview_only': False,
            'backup_original': True
        }
        if replacement_options:
            options.update(replacement_options)
        
        try:
            # HWPX 파일 읽기
            result = self.read_hwpx(hwpx_file)
            if not result['success']:
                return result
            
            # 참조 데이터 로드
            reference_dict = self.load_reference_data(reference_data)
            
            # 백업 생성
            if options['backup_original'] and not options['preview_only']:
                backup_path = Path(hwpx_file).with_suffix('.hwpx.backup')
                if not backup_path.exists():
                    shutil.copy2(hwpx_file, backup_path)
                    self.logger.info(f"백업 파일 생성: {backup_path}")
            
            # 텍스트 치환 수행
            replacement_results = []
            modified_text = result['text']
            total_replacements = 0
            
            # 진행률 표시
            items = reference_dict.items()
            if HAS_TQDM:
                items = tqdm(items, desc="텍스트 치환", unit="terms")
            
            for search_term, replacement_term in items:
                if not search_term.strip():
                    continue
                
                # 정규식 엔진 선택
                re_module = advanced_re if options['use_advanced_regex'] else re
                
                # 검색 패턴 생성
                if options['use_regex']:
                    pattern = search_term
                else:
                    escaped_term = re_module.escape(search_term)
                    if options['whole_word_only']:
                        pattern = r'\b' + escaped_term + r'\b'
                    else:
                        pattern = escaped_term
                
                # 검색 및 치환
                flags = 0 if options['case_sensitive'] else re_module.IGNORECASE
                
                if options['preview_only']:
                    matches = list(re_module.finditer(pattern, modified_text, flags))
                    if matches:
                        replacement_results.append({
                            'search_term': search_term,
                            'replacement_term': replacement_term,
                            'matches': len(matches),
                            'positions': [m.start() for m in matches]
                        })
                else:
                    if options['max_replacements_per_term'] > 0:
                        modified_text, count = re_module.subn(
                            pattern, replacement_term, modified_text, 
                            count=options['max_replacements_per_term'], flags=flags
                        )
                    else:
                        modified_text, count = re_module.subn(
                            pattern, replacement_term, modified_text, flags=flags
                        )
                    
                    total_replacements += count
                    
                    if count > 0:
                        replacement_results.append({
                            'search_term': search_term,
                            'replacement_term': replacement_term,
                            'count': count
                        })
                        self.logger.debug(f"'{search_term}' → '{replacement_term}' ({count}회)")
            
            # 처리 시간 계산
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            # 결과 구성
            replacement_summary = {
                'file_path': hwpx_file,
                'total_terms_processed': len(reference_dict),
                'total_replacements': total_replacements,
                'replacements': replacement_results,
                'preview_only': options['preview_only'],
                'processing_time': processing_time,
                'options_used': options,
                'success': True
            }
            
            if not options['preview_only']:
                replacement_summary['modified_text'] = modified_text
                replacement_summary['original_length'] = len(result['text'])
                replacement_summary['modified_length'] = len(modified_text)
            
            # 결과 파일 저장
            if output_file and not options['preview_only']:
                self._save_output_file(output_file, replacement_summary, modified_text)
            
            self.logger.info(f"처리 완료: {total_replacements}개 치환, {processing_time:.2f}초 소요")
            return replacement_summary
        
        except Exception as e:
            self.logger.error(f"처리 실패: {e}")
            return {
                'error': str(e),
                'success': False,
                'processing_time': (datetime.now() - start_time).total_seconds()
            }
    
    def _save_output_file(self, output_file: str, summary: Dict, modified_text: str):
        """출력 파일 저장"""
        try:
            output_path = Path(output_file)
            
            if output_path.suffix.lower() == '.txt':
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(modified_text)
            elif output_path.suffix.lower() == '.json':
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(summary, f, ensure_ascii=False, indent=2, default=str)
            
            summary['output_file'] = str(output_path)
            self.logger.info(f"결과 파일 저장: {output_path}")
            
        except Exception as e:
            summary['output_error'] = str(e)
            self.logger.error(f"출력 파일 저장 실패: {e}")
    
    def batch_process_folder(self, 
                           folder_path: str, 
                           reference_data: Union[str, Dict, List],
                           output_folder: Optional[str] = None,
                           file_pattern: str = "*.hwpx",
                           **kwargs) -> List[Dict]:
        """향상된 폴더 일괄 처리"""
        
        self.stats = ProcessingStats()
        self.stats.start_time = datetime.now()
        
        folder = Path(folder_path)
        results = []
        
        # 출력 폴더 설정
        if output_folder:
            output_dir = Path(output_folder)
        else:
            output_dir = folder / "processed"
        
        output_dir.mkdir(exist_ok=True)
        
        # 파일 목록 가져오기
        hwpx_files = list(folder.glob(file_pattern))
        self.stats.files_processed = len(hwpx_files)
        
        self.logger.info(f"일괄 처리 시작: {len(hwpx_files)}개 파일")
        
        # 진행률 표시
        file_iterator = hwpx_files
        if HAS_TQDM:
            file_iterator = tqdm(hwpx_files, desc="파일 처리", unit="files")
        
        for file_path in file_iterator:
            try:
                output_file = output_dir / f"{file_path.stem}_processed.txt"
                
                result = self.search_and_replace_text(
                    str(file_path), 
                    reference_data,
                    str(output_file),
                    **kwargs
                )
                
                result['source_file'] = str(file_path)
                results.append(result)
                
                if result['success']:
                    self.stats.files_successful += 1
                    self.stats.total_replacements += result.get('total_replacements', 0)
                    if not HAS_TQDM:  # TQDM이 없을 때만 개별 로그
                        self.logger.info(f"✓ {file_path.name}: {result['total_replacements']}개 치환")
                else:
                    self.stats.files_failed += 1
                    if not HAS_TQDM:
                        self.logger.error(f"✗ {file_path.name}: {result.get('error', '알 수 없는 오류')}")
                    
            except Exception as e:
                self.stats.files_failed += 1
                error_result = {
                    'source_file': str(file_path),
                    'success': False,
                    'error': str(e)
                }
                results.append(error_result)
                if not HAS_TQDM:
                    self.logger.error(f"✗ {file_path.name}: {str(e)}")
        
        # 통계 업데이트
        self.stats.end_time = datetime.now()
        self.stats.processing_time = (self.stats.end_time - self.stats.start_time).total_seconds()
        
        # 결과 요약 저장
        self._save_batch_summary(output_dir, results)
        
        self.logger.info(f"일괄 처리 완료: {self.stats.files_successful}/{self.stats.files_processed}개 성공")
        self.logger.info(f"총 처리 시간: {self.stats.processing_time:.2f}초")
        self.logger.info(f"총 치환 횟수: {self.stats.total_replacements}개")
        
        return results
    
    def _save_batch_summary(self, output_dir: Path, results: List[Dict]):
        """일괄 처리 결과 요약 저장"""
        summary_file = output_dir / "batch_processing_summary.json"
        
        summary_data = {
            'processing_stats': {
                'start_time': self.stats.start_time.isoformat() if self.stats.start_time else None,
                'end_time': self.stats.end_time.isoformat() if self.stats.end_time else None,
                'processing_time': self.stats.processing_time,
                'files_processed': self.stats.files_processed,
                'files_successful': self.stats.files_successful,
                'files_failed': self.stats.files_failed,
                'total_replacements': self.stats.total_replacements
            },
            'detailed_results': results
        }
        
        with open(summary_file, 'w', encoding='utf-8') as f:
            json.dump(summary_data, f, ensure_ascii=False, indent=2, default=str)
        
        self.logger.info(f"처리 요약 저장: {summary_file}")
    
    # 기존 메서드들 (간략화)
    def read_hwpx(self, file_path: str) -> Dict:
        """HWPX 파일 읽기"""
        try:
            self.temp_dir = tempfile.mkdtemp()
            
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)
            
            self._parse_document_info()
            self._parse_sections()
            
            full_text = self._extract_full_text()
            
            return {
                'text': full_text,
                'sections': self.sections,
                'metadata': {'file_path': file_path, 'char_count': len(full_text)},
                'success': True
            }
            
        except Exception as e:
            return {'error': str(e), 'success': False}
        finally:
            if self.temp_dir and Path(self.temp_dir).exists():
                shutil.rmtree(self.temp_dir)
    
    def _parse_document_info(self):
        """문서 정보 파싱"""
        try:
            doc_info_path = Path(self.temp_dir) / "DocInfo" / "document.xml"
            if doc_info_path.exists():
                # LXML 사용 가능하면 사용
                if HAS_LXML:
                    tree = LXML_ET.parse(str(doc_info_path))
                    root = tree.getroot()
                else:
                    tree = ET.parse(doc_info_path)
                    root = tree.getroot()
                
                summary_info = root.find('.//SUMMARYINFO')
                if summary_info is not None:
                    self.document_info = {
                        'title': summary_info.get('title', ''),
                        'author': summary_info.get('author', ''),
                        'date': summary_info.get('date', ''),
                    }
        except Exception as e:
            self.logger.debug(f"문서 정보 파싱 오류: {e}")
    
    def _parse_sections(self):
        """섹션 파싱"""
        try:
            bodytext_dir = Path(self.temp_dir) / "BodyText"
            if bodytext_dir.exists():
                section_files = sorted(bodytext_dir.glob("Section*.xml"))
                
                for i, section_file in enumerate(section_files):
                    if HAS_LXML:
                        tree = LXML_ET.parse(str(section_file))
                        root = tree.getroot()
                    else:
                        tree = ET.parse(section_file)
                        root = tree.getroot()
                    
                    text_content = self._extract_text_from_element(root)
                    
                    self.sections.append({
                        'section_index': i,
                        'text': text_content,
                        'file_name': section_file.name
                    })
        except Exception as e:
            self.logger.debug(f"섹션 파싱 오류: {e}")
    
    def _extract_text_from_element(self, element):
        """XML 요소에서 텍스트 추출"""
        text_parts = []
        
        def extract_recursive(elem):
            if elem.text:
                text_parts.append(elem.text)
            for child in elem:
                extract_recursive(child)
                if child.tail:
                    text_parts.append(child.tail)
        
        extract_recursive(element)
        full_text = ''.join(text_parts)
        return re.sub(r'\s+', ' ', full_text).strip()
    
    def _extract_full_text(self):
        """전체 텍스트 추출"""
        return '\n\n'.join(section['text'] for section in self.sections if section['text'])

# 사용 예시
if __name__ == "__main__":
    # 향상된 프로세서 생성
    processor = EnhancedHWPXProcessor(log_level="INFO")
    
    # 참조 데이터 (엑셀 지원)
    reference_data = {
        "구용어": "신용어",
        "HWPX": "한글문서",
        "변환": "치환"
    }
    
    # 단일 파일 처리 (향상된 옵션)
    result = processor.search_and_replace_text(
        hwpx_file="sample.hwpx",
        reference_data=reference_data,
        output_file="processed_sample.txt",
        replacement_options={
            'case_sensitive': False,
            'whole_word_only': True,
            'use_advanced_regex': True,  # 고급 정규식 엔진 사용
            'backup_original': True,     # 원본 백업
            'preview_only': False
        }
    )
    
    # 폴더 일괄 처리 (진행률 표시)
    batch_results = processor.batch_process_folder(
        folder_path="./documents",
        reference_data="reference.xlsx",  # 엑셀 파일 지원
        output_folder="./processed_documents"
    )
