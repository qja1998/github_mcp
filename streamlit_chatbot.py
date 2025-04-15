import streamlit as st
import asyncio
import shutil
import os
import dotenv
from collections import deque
import platform # platform 모듈 임포트

# --- Windows 환경일 경우 이벤트 루프 정책 설정 ---
if platform.system() == "Windows":
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception as e:
        st.warning(f"Failed to set WindowsProactorEventLoopPolicy: {e}")

# --- 초기 설정 및 환경 변수 로드 ---
try:
    dotenv.load_dotenv()
except Exception:
    pass # .env 파일 없어도 계속 진행

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4") # 기본값 설정

# --- 라이브러리 임포트 및 필수 변수 확인 ---
try:
    from agents import Agent, Runner, trace # trace는 선택적으로 사용
    from agents.mcp import MCPServer, MCPServerStdio
except ImportError:
    st.error("필수 라이브러리(`agents`, `agents-mcp`)를 찾을 수 없습니다.")
    st.stop()

if not GITHUB_PERSONAL_ACCESS_TOKEN:
    st.error("GITHUB_PERSONAL_ACCESS_TOKEN 환경 변수가 설정되지 않았습니다.")
    st.stop()

# --- 원본 run 함수의 로직을 분할한 비동기 함수들 ---

async def fetch_initial_repositories(github_path: str) -> str | None:
    """초기 레포지토리 목록을 가져옵니다 (원본 run 함수의 첫 부분)"""
    st.info(f"'{github_path}'에서 레포지토리 목록을 가져오는 중...")
    docker_env = f"GITHUB_PERSONAL_ACCESS_TOKEN={GITHUB_PERSONAL_ACCESS_TOKEN}"
    try:
        # 이 작업을 위해 별도의 서버 컨텍스트 생성
        async with MCPServerStdio(
            cache_tools_list=True, # 원본처럼 캐시 사용 (선택 사항)
            params={
                "command": "docker",
                "args": ["run", "-i", "--rm", "-e", docker_env, "ghcr.io/github/github-mcp-server"]
            }
        ) as server:
            # 레포지토리 목록 가져오기용 임시 Agent
            repo_agent = Agent(
                model=OPENAI_MODEL,
                name="RepoFetcher",
                instructions="List all repositories separated by '\\n'. Use the provided path.",
                mcp_servers=[server],
            )
            # 원본의 첫 Runner.run 호출과 유사
            result = await Runner.run(
                starting_agent=repo_agent,
                input=f"{github_path}의 모든 repository를 '\\n'로 구분해서 나열"
            )
            if hasattr(result, 'final_output'):
                st.success("레포지토리 목록 로딩 완료!")
                return result.final_output
            else:
                st.error("레포지토리 목록 결과 형식 오류")
                return None
    except Exception as e:
        st.error(f"레포지토리 목록 로딩 중 오류: {e}")
        import traceback
        print("--- fetch_initial_repositories ERROR ---")
        traceback.print_exc()
        print("--------------------------------------")
        return None

async def process_user_command(github_path: str, repositories: str, chat_history: deque, user_command: str) -> str:
    """사용자 명령(채팅 입력)을 처리합니다 (원본 run 함수의 루프 내부 로직)"""
    st.info("AI 응답 생성 중...")
    docker_env = f"GITHUB_PERSONAL_ACCESS_TOKEN={GITHUB_PERSONAL_ACCESS_TOKEN}"
    user_id = github_path.split("/")[-1] if "github.com" in github_path else "local_user"

    # 원본 PROMPT 구성 (상수 부분)
    PROMPT_BASE = f"""
    You are a potfolio assistant for '{user_id}'. You are given '{github_path}'.
    철저하게 {user_id} 관점에서 답변하라.
    You should answer questions about the git repository.
    NEVER answer questions about the directories that are not in the git repository.
    GET 호출을 할 때는 반드시 실제 있는 디렉토리를 사용해야 한다.
    답변은 절대 지어내지 말고, 항상 사실에 기반해야 한다.
    You should answer questions about the git repository.
    repositories:
    {repositories}
    """

    # 원본 END_PROMPT 구성 (동적 부분)
    history_str = "\n".join([f"{item['role']}: {item['content']}" for turn in chat_history for item in turn])
    END_PROMPT = f"""
    history: {history_str}
    user: {user_command}"""

    full_prompt = PROMPT_BASE + END_PROMPT
    # print(f"--- Full Prompt ---\n{full_prompt}\n-------------------") # 디버깅용

    try:
        # 이 상호작용을 위해 별도의 서버 컨텍스트 생성
        async with MCPServerStdio(
            cache_tools_list=True,
            params={
                "command": "docker",
                "args": ["run", "-i", "--rm", "-e", docker_env, "ghcr.io/github/github-mcp-server"]
            }
        ) as server:
            # 채팅 응답용 Agent (원본 run 함수의 agent와 동일)
            chat_agent = Agent(
                model=OPENAI_MODEL,
                name="Assistant",
                instructions="Answer questions about the git repositories based on the provided context and history.", # 약간 수정
                mcp_servers=[server],
            )

            # 원본의 루프 내 Runner.run 호출과 유사
            # trace 사용 가능: with trace(workflow_name="Streamlit Chat Interaction"):
            result = await Runner.run(starting_agent=chat_agent, input=full_prompt)

            if hasattr(result, 'final_output'):
                 st.success("AI 응답 생성 완료!")
                 return result.final_output
            else:
                st.error("AI 응답 결과 형식 오류")
                return "오류: 응답 형식이 올바르지 않습니다."

    except Exception as e:
        st.error(f"AI 응답 생성 중 오류: {e}")
        import traceback
        print("--- process_user_command ERROR ---")
        traceback.print_exc()
        print("----------------------------------")
        return "오류: 응답 생성에 실패했습니다."


# --- Streamlit UI 및 상태 관리 ---

st.set_page_config(page_title="GitHub Repo Chat", layout="wide")
st.title("📁 GitHub Repository Chat")

# 세션 상태 초기화 (원본 구조 반영)
if "github_path" not in st.session_state:
    st.session_state.github_path = ""
if "user_id" not in st.session_state:
    st.session_state.user_id = ""
if "repositories" not in st.session_state: # 레포 목록 문자열 (성공 시) 또는 None (실패/미로드)
    st.session_state.repositories = None
if "chat_history_deque" not in st.session_state:
    st.session_state.chat_history_deque = deque([], maxlen=5) # 원본과 동일하게 deque 사용
if "initial_fetch_done" not in st.session_state:
     st.session_state.initial_fetch_done = False # 초기 로딩 완료 여부 플래그


# --- 사이드바: GitHub 경로 입력 ---
with st.sidebar:
    st.header("Target Repository")
    new_path = st.text_input("GitHub User/Org URL:", value=st.session_state.github_path)

    # 경로 변경 시 상태 초기화 및 초기 정보 로드 트리거
    if new_path and new_path != st.session_state.github_path:
        st.session_state.github_path = new_path
        st.session_state.user_id = new_path.split("/")[-1] if "github.com" in new_path else "local_user"
        st.session_state.repositories = None # 레포 목록 초기화
        st.session_state.chat_history_deque = deque([], maxlen=5) # 대화 기록 초기화
        st.session_state.initial_fetch_done = False # 초기 로딩 플래그 리셋
        # 초기 레포지토리 목록 로드 실행
        with st.spinner(f"Loading repositories from {new_path}..."):
            repos = asyncio.run(fetch_initial_repositories(new_path))
            st.session_state.repositories = repos # 결과 저장
            st.session_state.initial_fetch_done = True # 로딩 시도 완료
        st.rerun() # 상태 저장 후 UI 즉시 새로고침

    # 로드된 레포지토리 목록 표시
    if st.session_state.repositories:
        st.subheader("Loaded Repositories")
        with st.expander("View List"):
            st.text_area("", st.session_state.repositories, height=150, disabled=True)
    elif st.session_state.initial_fetch_done and st.session_state.github_path:
        st.warning("Failed to load repositories. Check path or logs.")


# --- 메인 영역: 채팅 인터페이스 ---
st.header("Chat")

# 이전 대화 내용 표시 (deque 사용)
# deque는 순서가 중요하므로, 저장된 순서대로 표시
if st.session_state.github_path and st.session_state.repositories:
    for turn in st.session_state.chat_history_deque:
        # turn은 [{'role':'user', 'content':'...'}, {'role':'assistant', 'content':'...'}] 형태
        if len(turn) >= 1 and turn[0]['role'] == 'user':
             with st.chat_message("user"):
                 st.markdown(turn[0]['content'])
        if len(turn) == 2 and turn[1]['role'] == 'assistant':
             with st.chat_message("assistant"):
                 st.markdown(turn[1]['content'])

# 채팅 입력 처리
if st.session_state.github_path and st.session_state.repositories:
    if user_command := st.chat_input("Ask about the repositories..."):
        # 사용자 입력 표시 (deque 업데이트 전에)
        with st.chat_message("user"):
            st.markdown(user_command)

        # 비동기 함수 호출하여 응답 받기
        ai_response = asyncio.run(
            process_user_command(
                st.session_state.github_path,
                st.session_state.repositories,
                st.session_state.chat_history_deque, # 현재까지의 deque 전달
                user_command
            )
        )

        # AI 응답 표시
        with st.chat_message("assistant"):
            st.markdown(ai_response)

        # 대화 기록 업데이트 (deque) - 원본 구조 유지
        current_chat_turn = [
            {"role": "user", "content": user_command},
            {"role": "assistant", "content": ai_response},
        ]
        st.session_state.chat_history_deque.append(current_chat_turn)

        # Rerun 필요 없음: Streamlit이 chat_input 처리 후 자동으로 rerun함
        # 단, deque 업데이트를 즉시 반영하려면 필요할 수 있으나, 다음 입력 시 반영됨.
        # st.rerun()

elif not st.session_state.github_path:
    st.info("Please enter a GitHub User/Org URL in the sidebar to start.")
elif st.session_state.initial_fetch_done and not st.session_state.repositories:
     st.warning("Repository loading failed. Cannot start chat.")
elif not st.session_state.initial_fetch_done and st.session_state.github_path:
     st.info("Loading repository information...")


# --- 원본 스크립트의 main 가드 부분 (Streamlit에서는 불필요) ---
# if __name__ == "__main__":
#     # uvx check (Streamlit 환경에서는 다른 방식으로 관리)
#     # asyncio.run(main()) # Streamlit이 실행 흐름을 관리