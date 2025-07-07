#!/usr/bin/env python3
"""
HWPX 텍스트 처리 CLI 도구

사용법:
    python hwpx_processor.py filename.hwpx
    python hwpx_processor.py filename.hwpx -r reference.csv
    python hwpx_processor.py -b ./documents -r terms.xlsx
"""

import argparse
import sys
from pathlib import Path
import json
from datetime import datetime

# 기존 EnhancedHWPXProcessor 임포트
from enhanced_hwpx_processor import EnhancedHWPXProcessor

def create_sample_reference_files():
    """샘플 참조 데이터 파일들 생성"""
    
    # 샘플 CSV 파일
    csv_content = """검색어,치환어
HWPX,한글문서파일
변환,치환
검색,탐색
구용어,신용어
클라우드,구름컴퓨팅
빅데이터,대용량데이터
AI,인공지능
IoT,사물인터넷
머신러닝,기계학습"""
    
    with open('sample_terms.csv', 'w', encoding='utf-8') as f:
        f.write(csv_content)
    
    # 샘플 JSON 파일
    json_content = {
        "Old Brand": "New Brand",
        "구브랜드": "신브랜드", 
        "legacy": "modern",
        "deprecated": "updated",
        "obsolete": "current"
    }
    
    with open('sample_terms.json', 'w', encoding='utf-8') as f:
        json.dump(json_content, f, ensure_ascii=False, indent=2)
    
    print("📁 샘플 참조 파일 생성 완료:")
    print("  - sample_terms.csv")
    print("  - sample_terms.json")

def validate_file_exists(file_path: str, file_type: str = "파일") -> Path:
    """파일 존재 여부 확인"""
    path = Path(file_path)
    if not path.exists():
        print(f"❌ 오류: {file_type}을(를) 찾을 수 없습니다: {file_path}")
        sys.exit(1)
    return path

def get_output_path(input_file: Path, output_dir: Path = None, suffix: str = "_processed") -> Path:
    """출력 파일 경로 생성"""
    if output_dir:
        output_dir.mkdir(exist_ok=True)
        return output_dir / f"{input_file.stem}{suffix}.txt"
    else:
        return input_file.parent / f"{input_file.stem}{suffix}.txt"

def print_processing_result(result: dict, file_name: str):
    """처리 결과 출력"""
    if result['success']:
        print(f"✅ {file_name} 처리 완료!")
        print(f"   📊 총 {result['total_replacements']}개 항목이 치환되었습니다.")
        print(f"   ⏱️  처리 시간: {result.get('processing_time', 0):.2f}초")
        
        if result['replacements']:
            print("   📝 치환 내역:")
            for replacement in result['replacements'][:10]:  # 최대 10개만 표시
                if isinstance(replacement, dict):
                    search_term = replacement['search_term']
                    replace_term = replacement['replacement_term']
                    count = replacement['count']
                    print(f"      '{search_term}' → '{replace_term}' ({count}회)")
            
            if len(result['replacements']) > 10:
                print(f"      ... 및 {len(result['replacements']) - 10}개 더")
        
        if 'output_file' in result:
            print(f"   📄 결과 파일: {result['output_file']}")
    else:
        print(f"❌ {file_name} 처리 실패!")
        print(f"   오류: {result.get('error', '알 수 없는 오류')}")

def main():
    parser = argparse.ArgumentParser(
        description='HWPX 파일의 텍스트를 검색하고 참조 데이터로 치환합니다.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python %(prog)s document.hwpx
  python %(prog)s document.hwpx -r terms.csv
  python %(prog)s document.hwpx -r terms.json -o output.txt
  python %(prog)s -b ./documents -r terms.xlsx
  python %(prog)s --create-samples
  python %(prog)s document.hwpx --preview
        """
    )
    
    # 위치 인수 (파일 또는 폴더)
    parser.add_argument(
        'input_path', 
        nargs='?', 
        help='처리할 HWPX 파일 또는 폴더 경로'
    )
    
    # 옵션 인수들
    parser.add_argument(
        '-r', '--reference', 
        help='참조 데이터 파일 (CSV, JSON, XLSX)',
        default='sample_terms.csv'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='출력 파일 경로 (단일 파일 처리시만)'
    )
    
    parser.add_argument(
        '-b', '--batch',
        help='폴더 일괄 처리 (input_path 대신 사용)',
        metavar='FOLDER_PATH'
    )
    
    parser.add_argument(
        '--output-dir',
        help='출력 디렉토리 (일괄 처리시)',
        default='./processed'
    )
    
    parser.add_argument(
        '--preview',
        action='store_true',
        help='실제 변경하지 않고 미리보기만'
    )
    
    parser.add_argument(
        '--case-sensitive',
        action='store_true',
        help='대소문자 구분'
    )
    
    parser.add_argument(
        '--whole-word',
        action='store_true',
        help='전체 단어만 매칭'
    )
    
    parser.add_argument(
        '--no-backup',
        action='store_true',
        help='원본 파일 백업 안함'
    )
    
    parser.add_argument(
        '--max-replacements',
        type=int,
        default=-1,
        help='용어당 최대 치환 횟수 (-1: 무제한)'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='로그 레벨'
    )
    
    parser.add_argument(
        '--create-samples',
        action='store_true',
        help='샘플 참조 데이터 파일 생성'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='HWPX Processor v1.0.0'
    )
    
    args = parser.parse_args()
    
    # 샘플 파일 생성
    if args.create_samples:
        create_sample_reference_files()
        return
    
    # 입력 검증
    if not args.input_path and not args.batch:
        parser.error("HWPX 파일 경로 또는 --batch 옵션이 필요합니다.")
    
    # 프로세서 초기화
    print(f"🚀 HWPX 텍스트 처리기 시작 (로그 레벨: {args.log_level})")
    processor = EnhancedHWPXProcessor(log_level=args.log_level)
    
    # 참조 데이터 파일 확인
    if Path(args.reference).exists():
        print(f"📋 참조 데이터: {args.reference}")
        reference_data = args.reference
    else:
        print(f"⚠️  참조 데이터 파일이 없습니다: {args.reference}")
        print("   sample_terms.csv 파일을 생성하려면 --create-samples 옵션을 사용하세요.")
        
        # 기본 참조 데이터 사용
        reference_data = {
            "예시용어": "변경된용어",
            "HWPX": "한글문서"
        }
        print("   기본 참조 데이터를 사용합니다.")
    
    # 치환 옵션 설정
    replacement_options = {
        'case_sensitive': args.case_sensitive,
        'whole_word_only': args.whole_word,
        'preview_only': args.preview,
        'backup_original': not args.no_backup,
        'max_replacements_per_term': args.max_replacements,
        'use_advanced_regex': True
    }
    
    # 일괄 처리
    if args.batch:
        batch_folder = validate_file_exists(args.batch, "폴더")
        print(f"📁 폴더 일괄 처리: {batch_folder}")
        
        results = processor.batch_process_folder(
            folder_path=str(batch_folder),
            reference_data=reference_data,
            output_folder=args.output_dir,
            replacement_options=replacement_options
        )
        
        # 결과 요약
        successful = len([r for r in results if r.get('success')])
        total_replacements = sum(r.get('total_replacements', 0) for r in results)
        
        print(f"\n📊 일괄 처리 완료!")
        print(f"   성공: {successful}/{len(results)}개 파일")
        print(f"   총 치환: {total_replacements}개")
        print(f"   결과 폴더: {args.output_dir}")
        
    # 단일 파일 처리
    else:
        input_file = validate_file_exists(args.input_path, "HWPX 파일")
        
        if input_file.suffix.lower() != '.hwpx':
            print(f"⚠️  경고: {input_file.name}은 HWPX 파일이 아닐 수 있습니다.")
        
        # 출력 파일 경로 결정
        if args.output:
            output_file = args.output
        else:
            output_file = get_output_path(input_file)
        
        print(f"📄 파일 처리: {input_file.name}")
        if args.preview:
            print("👀 미리보기 모드 (실제 변경하지 않음)")
        else:
            print(f"💾 결과 저장: {output_file}")
        
        # 처리 실행
        result = processor.search_and_replace_text(
            hwpx_file=str(input_file),
            reference_data=reference_data,
            output_file=str(output_file) if not args.preview else None,
            replacement_options=replacement_options
        )
        
        # 결과 출력
        print_processing_result(result, input_file.name)
        
        # 미리보기 결과 상세 출력
        if args.preview and result.get('success') and result.get('replacements'):
            print(f"\n🔍 미리보기 상세:")
            for replacement in result['replacements'][:5]:  # 최대 5개
                if isinstance(replacement, dict) and 'positions' in replacement:
                    search_term = replacement['search_term']
                    positions = replacement['positions'][:3]  # 최대 3개 위치
                    print(f"   '{search_term}' 발견 위치: {positions}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  사용자에 의해 중단되었습니다.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류 발생: {e}")
        sys.exit(1)
