import argparse
import sys
from somers_agent import SomersAgent

def main():
    parser = argparse.ArgumentParser(description="소머즈 에이전트 전문 문서 검색 및 일일 AI 학습기 CLI")
    parser.add_argument("topic", type=str, help="검색 및 학습할 주제 (예: '식품공장 위생배관', 'MES 견적')")
    parser.add_argument("--docx", type=str, default="학습결과_보고서.docx", help="저장할 Word 보고서 파일명")
    parser.add_argument("--md", type=str, default="학습결과_보고서.md", help="저장할 마크다운 보고서 파일명")
    parser.add_argument("--chart", type=str, default="신뢰도_분포.png", help="저장할 시각화 차트 파일명")
    parser.add_argument("--sync", action="store_true", help="학습 완료 후 깃허브 자동 동기화 활성화")
    parser.add_argument("--repo", type=str, default="https://github.com/pjw1356-hub/connected-AI.git", help="동기화할 깃허브 저장소 주소")
    
    args = parser.parse_args()
    
    agent = SomersAgent()
    try:
        # 1. 일일 학습 프로세스 작동 (수집, 평가, knowledge_base.json 저장 누적)
        learnings = agent.run_daily_learning(args.topic)
        
        # 2. 산출물 파일 생성
        agent.generate_docx_report(learnings, args.docx)
        agent.generate_md_report(learnings, args.md)
        agent.visualize_statistics(learnings, args.chart)
        
        # 3. 깃허브 동기화 진행
        if args.sync:
            agent.sync_to_github(args.repo)
        
        print("\n" + "="*50)
        print("★ 소머즈 에이전트 일일 지식 학습 성공 ★")
        print(f"1. AI 지식베이스 누적 저장 완료: knowledge_base.json")
        print(f"2. 마크다운 보고서 생성 완료: {args.md}")
        print(f"3. MS Word 보고서 생성 완료: {args.docx}")
        print(f"4. 신뢰도 평가 통계 차트 완료: {args.chart}")
        if args.sync:
            print(f"5. 깃허브 저장소 동기화 시도 완료")
        print("="*50)
        
    except Exception as e:
        print(f"오류가 발생했습니다: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
