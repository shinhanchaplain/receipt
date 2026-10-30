import os
import sys
import subprocess

# 현재 스크립트가 위치한 폴더로 작업 경로 이동 (한글 경로 완벽 지원)
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# 콘솔 UTF-8 출력 설정
if sys.stdout and sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

print("=" * 60)
print("  [ GitHub 업로드: shinhanchaplain/receipt ]")
print("=" * 60)
print(f"작업 폴더: {script_dir}\n")

repo_url = "https://github.com/shinhanchaplain/receipt.git"

def run_git_cmd(args, desc):
    print(f"▶ {desc}...")
    result = subprocess.run(args, capture_output=False)
    return result.returncode

# 1. 기존 origin 제거 (있을 경우)
subprocess.run(["git", "remote", "remove", "origin"], capture_output=True)

# 2. 새로운 origin 등록
code1 = run_git_cmd(["git", "remote", "add", "origin", repo_url], "1. GitHub 원격 저장소 연결")

# 3. main 브랜치 설정
code2 = run_git_cmd(["git", "branch", "-M", "main"], "2. 기본 브랜치(main) 설정")

# 4. GitHub에 푸시
print("\n▶ 3. GitHub에 파일 푸시 중 (잠시만 기다려주세요)...")
print("※ 브라우저 로그인 창이 뜨면 로그인/승인(Authorize)을 클릭해 주세요.\n")
push_result = subprocess.run(["git", "push", "-u", "origin", "main"])

print("\n" + "=" * 60)
if push_result.returncode == 0:
    print("🎉 [성공] GitHub에 모든 파일이 정상적으로 업로드되었습니다!")
    print("이제 브라우저에서 아래 페이지를 열고 배포를 진행하세요:")
    print("1. GitHub 확인: https://github.com/shinhanchaplain/receipt")
    print("2. Streamlit 배포: https://share.streamlit.io/")
else:
    print("⚠️ [안내] 푸시 도중 로그인이 필요하거나 중단되었습니다.")
    print("GitHub 권한 또는 인터넷 연결 상태를 확인 후 다시 실행해주세요.")
print("=" * 60 + "\n")
