import argparse
import sys
import datetime
from somers_agent import SomersAgent

def sanitize_input(text: str) -> str:
    try:
        # 유니코드 surrogates 문자를 안전하게 무시 및 치환
        return text.encode('utf-8', 'surrogateescape').decode('utf-8', 'ignore')
    except Exception:
        return text

def get_interactive_inputs():
    """
    사용자로부터 매 학습 실행 전에 세부 조건들을 대화식으로 입력받습니다.
    """
    print("\n" + "="*60)
    print("★ 소머즈 에이전트(SomersAgent) 일일 학습 조건 사전 설정 ★")
    print("="*60)
    
    # 1. 학습 주제 입력
    topic_raw = input("1. 오늘 학습할 전문분야 주제를 입력하세요 (예: 식품공장 위생배관): ").strip()
    topic = sanitize_input(topic_raw)
    while not topic:
        topic_raw = input("  -> 주제는 필수 입력입니다. 다시 입력하세요: ").strip()
        topic = sanitize_input(topic_raw)
        
    # 2. 최신성 연도 제한
    max_age_str = input("2. 허용할 문서의 최신성 기준 연도를 입력하세요 (기본값: 5년 이내): ").strip()
    if not max_age_str:
        max_age = 5
    else:
        try:
            max_age = int(max_age_str)
        except ValueError:
            print("  -> 올바른 숫자가 아니어서 기본값(5년)으로 설정합니다.")
            max_age = 5
            
    # 3. 추가 세부 요청사항 (제약조건)
    user_request_raw = input("3. 학습 및 검토 시 반영할 추가 상세 조건/요청사항을 입력하세요\n   (예: '수치 중심 설계 규격 검토', '국가 표준 위주 수집'): ").strip()
    user_request = sanitize_input(user_request_raw)
    
    # 4. 깃허브 동기화 여부
    sync_yn = input("4. 학습 완료 후 결과를 깃허브 저장소로 바로 동기화할까요? (y/N): ").strip().lower()
    sync = sync_yn in ["y", "yes"]
    
    print("="*60 + "\n")
    
    return topic, {
        "max_age": max_age,
        "user_request": user_request
    }, sync

def main():
    parser = argparse.ArgumentParser(description="소머즈 에이전트 전문 문서 검색 및 일일 AI 학습기 CLI")
    parser.add_argument("topic", type=str, nargs="?", default=None, help="검색 및 학습할 주제 (생략 시 대화형 모드 구동)")
    parser.add_argument("--docx", type=str, default="학습결과_보고서.docx", help="저장할 Word 보고서 파일명")
    parser.add_argument("--md", type=str, default="학습결과_보고서.md", help="저장할 마크다운 보고서 파일명")
    parser.add_argument("--chart", type=str, default="신뢰도_분포.png", help="저장할 시각화 차트 파일명")
    parser.add_argument("--sync", action="store_true", help="학습 완료 후 깃허브 자동 동기화 활성화")
    parser.add_argument("--repo", type=str, default="https://github.com/pjw1356-hub/connected-AI.git", help="동기화할 깃허브 저장소 주소")
    
    args = parser.parse_args()
    
    # 주제가 아규먼트로 제공되지 않은 경우, 인터랙티브 대화형 입력 모드 구동
    if args.topic is None:
        topic, constraints, sync = get_interactive_inputs()
    else:
        topic = args.topic
        constraints = {
            "max_age": 5,
            "user_request": ""
        }
        sync = args.sync
        
    agent = SomersAgent()
    try:
        # 1. 일일 학습 프로세스 작동 (수집, 평가, 사용자 조건 및 지식베이스 누적)
        learnings = agent.run_daily_learning(topic, constraints)
        
        # 2. 산출물 파일 생성
        agent.generate_docx_report(learnings, args.docx)
        agent.generate_md_report(learnings, args.md)
        agent.visualize_statistics(learnings, args.chart)
        
        # 3. 깃허브 동기화 진행
        if sync:
            agent.sync_to_github(args.repo)
        
        print("\n" + "="*50)
        print("★ 소머즈 에이전트 일일 지식 학습 성공 ★")
        print(f"1. AI 지식베이스 누적 저장 완료: knowledge_base.json")
        print(f"2. 마크다운 보고서 생성 완료: {args.md}")
        print(f"3. MS Word 보고서 생성 완료: {args.docx}")
        print(f"4. 신뢰도 평가 통계 차트 완료: {args.chart}")
        if sync:
            print(f"5. 깃허브 저장소 동기화 완료")
        print("="*50)
        
    except Exception as e:
        print(f"오류가 발생했습니다: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
