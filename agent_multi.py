import asyncio
import shutil
import os
import json  # JSON 파싱을 위해 추가
import dotenv

dotenv.load_dotenv(override=True)
ROOT = os.path.abspath(os.path.dirname(__file__))

from collections import deque

# GITHUB_PERSONAL_ACCESS_TOKEN는 더 이상 서버 시작 시 필수가 아님
# GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

from agents import (
    Agent,
    Runner,
    trace,
)  # 정확한 import 경로는 라이브러리에 따라 다를 수 있음
from agents.mcp import MCPServer, MCPServerStdio  # 정확한 import 경로

# 사용자별 세션 ID를 저장하기 위한 간단한 딕셔너리 (실제 서비스에서는 더 견고한 저장 방식 필요)
user_sessions = {}


# 새로운 GitHub MCP 서버 실행 함수 (세션 기반)
async def start_github_mcp_server_session_based():
    # 서버 시작 시 GITHUB_PERSONAL_ACCESS_TOKEN 환경 변수는 더 이상 필수가 아님.
    # 서버는 기본 GHE 호스트 또는 공용 GitHub를 사용하도록 설정됨.
    # (예: docker run ... ghcr.io/github/github-mcp-server stdio --gh-host "https:#api.github.com")
    # MCPServerStdio의 params에서 env의 GITHUB_PERSONAL_ACCESS_TOKEN 제거
    mcp_server_process = MCPServerStdio(
        name="GithubMCPMultiUser",
        cache_tools_list=True,
        params={
            "command": "C:/Users/kwon/Desktop/repo/github-mcp-server/cmd/github-mcp-server",  # 로컬 빌드된 서버 실행 파일 경로
            "args": [
                "stdio",
                "--gh-host",
                "https:#api.github.com",  # 기본 호스트
                "--log-file",
                "server.log",
                "--enable-command-logging",
                # GITHUB_TOOLSETS, GITHUB_DYNAMIC_TOOLSETS 등 필요에 따라 추가
            ],
            "env": {
                # "GITHUB_PERSONAL_ACCESS_TOKEN": GITHUB_PERSONAL_ACCESS_TOKEN, # 제거 또는 서버 관리용 기본 토큰(다른 용도)
            },
        },
    )
    return mcp_server_process


# 사용자를 위한 토큰 등록 및 세션 ID 획득 함수
async def register_user_token(
    mcp_server: MCPServer,
    user_id: str,
    github_pat: str,
    github_host: str = None,
    client_info: dict = None,
):
    if not github_pat:
        print(f"Error: GitHub PAT is required for user {user_id}.")
        return None

    tool_name = "register_session_token"
    arguments = {"personal_access_token": github_pat}
    if github_host:
        arguments["github_host"] = github_host

    meta = {}
    if client_info:
        meta["clientInfo"] = client_info

    print(f"Attempting to register token for user: {user_id} with meta: {meta}")

    try:
        # `agents` 라이브러리가 `call_tool`에 `meta` 파라미터를 지원한다고 가정
        # 지원하지 않는다면, 라이브러리 문서를 확인하거나 직접 JSON-RPC 메시지를 구성하여 전송해야 함.
        # response = await mcp_server.call_tool(tool_name, arguments, meta=meta) # 가상의 meta 파라미터

        # MCPServerStdio의 call_tool_with_meta 와 같은 명시적 함수가 필요할 수 있습니다.
        # 또는, Agent 클래스나 Runner를 통해 meta 정보를 전달할 수 있는지 확인해야 합니다.
        # 아래는 agents.mcp.MCPServer 클래스에 send_request 와 같은 저수준 메소드가 있고,
        # 이를 통해 _meta를 포함한 전체 JSON-RPC 요청을 보낼 수 있다고 가정한 예시입니다.
        # 실제 사용법은 agents 라이브러리 문서를 참조해야 합니다.

        # 가상의 send_json_rpc_request 함수 (실제 라이브러리 기능에 따라 구현 필요)
        request_id = f"reg-{user_id}-{os.urandom(4).hex()}"
        rpc_request = {
            "jsonrpc": "2.0",
            "id": request_id,
            "method": "tools/call",
            "params": {"name": tool_name, "arguments": arguments},
        }
        if meta:
            rpc_request["_meta"] = meta

        # MCPServerStdio 인스턴스가 직접 JSON RPC 요청을 보내고 응답을 받는 메소드를 제공해야 함.
        # 예를 들어 response_data = await mcp_server.send_request(rpc_request)
        # 아래는 단순화를 위해 Agent를 통한 호출을 시도하지만, _meta 전달이 문제될 수 있습니다.

        # Agent를 사용하여 도구 호출 시 _meta 전달 방법 확인 필요
        # 임시 Agent 생성 또는 mcp_server 객체에서 직접 호출 가능한 메소드 탐색
        temp_agent_for_reg = Agent(name="TokenRegistrar", mcp_servers=[mcp_server])
        # Agent 클래스가 call_tool 메소드에 meta 인자를 지원하는지 확인 필요
        response = await temp_agent_for_reg.call_tool(
            tool_name, arguments
        )  # TODO: _meta 전달 방법 확인!

        if response and response.content:
            # 응답이 {"session_id": "xxxx"} 형태의 JSON 문자열을 포함하는 text content로 올 것임
            text_content = response.content[0].get("text", "")
            session_data = json.loads(text_content)
            session_id = session_data.get("session_id")
            if session_id:
                user_sessions[user_id] = session_id  # 사용자 ID와 세션 ID 매핑 저장
                print(f"Token registered for user {user_id}. Session ID: {session_id}")
                return session_id
            else:
                print(
                    f"Error registering token for user {user_id}: 'session_id' not in response. Full response: {text_content}"
                )
        else:
            print(
                f"Error registering token for user {user_id}: No valid response content. Response: {response}"
            )

    except Exception as e:
        print(f"Exception during token registration for user {user_id}: {e}")
    return None


async def run_user_portfolio_assistant(
    mcp_server: MCPServer, user_id: str, user_github_url: str, user_pat: str
):
    session_id = user_sessions.get(user_id)
    if not session_id:
        # PAT를 사용하여 세션 ID 등록 시도
        print(f"No active session for {user_id}. Attempting to register token...")
        client_info_for_reg = {"name": "PortfolioAssistPythonClient", "version": "1.1"}
        session_id = await register_user_token(
            mcp_server, user_id, user_pat, client_info=client_info_for_reg
        )
        if not session_id:
            print(f"Failed to establish session for {user_id}. Aborting.")
            return

    # Agent 생성 시 mcp_servers 리스트 전달
    # Agent 또는 Runner가 도구 호출 시 _meta 필드를 전달할 수 있는 방법을 찾아야 함.
    # 만약 Agent 생성 시 또는 Runner.run 호출 시 meta 정보를 전역적으로 설정할 수 있다면 그 방법을 사용.
    # 그렇지 않다면, 각 tool_call 에 meta를 삽입할 수 있도록 agents 라이브러리 사용법을 확인해야 함.

    # Agent가 모든 요청에 특정 meta를 포함하도록 설정할 수 있다면 가장 좋음
    # 예: agent = Agent(..., default_meta={"session_id": session_id, "clientInfo": ...}) (가상)
    agent = Agent(
        model=OPENAI_MODEL,  # 귀하의 OPENAI_MODEL 변수 사용
        name=f"Assistant_{user_id}",
        instructions=f"Answer questions about the git repositories for {user_id}. Use the provided tools.",
        mcp_servers=[mcp_server],
        # default_tool_call_meta = {"session_id": session_id} # 만약 라이브러리가 이런 기능을 지원한다면
    )

    # 초기 리포지토리 목록 가져오기 - 이 호출에도 session_id가 필요
    # Runner.run 호출 시 meta 전달 방법 확인
    initial_repositories_prompt = (
        f"List all repositories for the user at {user_github_url}, separated by '\\n'."
    )
    print(f"Fetching initial repositories for {user_id}...")

    # Runner.run 호출 시 _meta 전달
    # 이 부분은 agents 라이브러리가 어떻게 _meta 필드 전송을 지원하는지에 따라 크게 달라집니다.
    # 라이브러리가 Runner.run(..., meta=...) 또는 Agent(..., default_meta=...) 등을 지원해야 합니다.
    # 아래는 meta를 지원한다고 가정한 예시입니다. 실제 API를 확인하세요.

    # 현재 agents 라이브러리 구조상 Runner.run에 직접 meta를 넘기기 어려울 수 있습니다.
    # Agent 객체가 생성될 때 meta 정보를 설정하거나,
    # MCP Server 객체를 통해 저수준으로 요청을 보내야 할 수 있습니다.
    meta_for_runner = {
        "session_id": session_id,
        "clientInfo": {
            "name": "PortfolioAssistPythonClient",
            "version": "1.1",
        },  # 필요시 전달
    }

    # repositories_result = await Runner.run(starting_agent=agent, input=initial_repositories_prompt, meta=meta_for_runner) # 가상 meta 파라미터
    # 위 `meta` 파라미터는 `agents` 라이브러리에 실제 존재해야 합니다.
    # 만약 없다면, `agent.call_tool`을 직접 사용하거나, `mcp_server` 객체를 통해
    # `send_request`와 같은 저수준 API를 사용하여 `_meta`를 포함한 JSON-RPC 요청을 만들어야 합니다.

    # 임시로, Agent가 생성될 때 meta 정보가 설정되었다고 가정하고 진행합니다.
    # 또는, get_client 함수가 호출될 때 Agent가 어떤식으로든 session_id를 MCP 서버에 전달해야 합니다.
    # 가장 확실한 방법은 mcp_server.call_tool을 직접 호출하는 것입니다.
    # (단, Agent 프레임워크의 추상화를 우회하게 될 수 있습니다)

    # 예시: 첫 번째 호출을 위한 _meta 설정 (agents 라이브러리 지원 필요)
    # Agent의 다음 호출에 사용될 meta를 설정하는 메서드가 있다면 사용
    # agent.set_next_call_meta(meta_for_runner)

    print(
        f"DEBUG: Using session_id: {session_id} for user {user_id} for initial repo list"
    )
    # repositories_result = await agent.run(input=initial_repositories_prompt) # agent.run이 meta를 내부적으로 처리한다면
    # 또는 저수준 호출:
    try:
        # 이 부분은 agents 라이브러리의 MCP 통신 방식에 따라 크게 달라집니다.
        # 가장 좋은 방법은 MCPServerStdio 또는 MCPServer 클래스에
        # send_request(method_name, params, meta_data) 와 같은 메소드가 있어 직접 호출하는 것입니다.
        # 여기서는 Runner.run을 사용하되, meta 전달이 가능하다고 가정합니다.
        # 하지만 현재 제공된 코드에는 Runner.run에 meta를 전달하는 부분이 명확하지 않으므로,
        # Agent가 mcp_servers를 통해 통신할 때 meta를 잘 전달하도록 Agent 또는 MCPServer 클래스 수정/확장이 필요할 수 있습니다.

        # 우선은 Agent가 내부적으로 meta를 잘 처리한다고 가정하고 진행합니다.
        # (실제로는 이 부분이 가장 큰 변경 지점일 수 있습니다)
        # Agent 클래스가 생성 시점에 mcp_server에 대한 호출 시 사용할 meta 정보를 설정할 수 있어야 합니다.
        # 예: agent = Agent(..., mcp_server_call_options={"meta": meta_for_runner})

        # 만약 Agent가 자동으로 meta를 처리하지 않는다면, 직접 tool을 호출해야 합니다.
        # 예시: list_repos_tool_response = await agent.call_tool("list_repositories", {"owner": user_id_from_url}, meta=meta_for_runner)
        # 실제 도구 이름과 파라미터는 MCP 서버에 정의된 대로 사용해야 합니다.
        # 지금은 "list_repositories"가 없으므로, search_repositories 사용 예시
        search_query = f"user:{user_github_url.split('/')[-1]} fork:true"  # 예시 쿼리
        # repositories_result = await agent.call_tool(
        #     "search_repositories",
        #     {"query": search_query},
        #     # meta=meta_for_runner # call_tool이 meta를 지원해야 함
        # )
        # print(f"Repositories result: {repositories_result}")
        # repositories_output = repositories_result.final_output if repositories_result else "Could not fetch repositories."
        # 현재 코드에서는 repositories = await Runner.run(...) 이 있으므로, 이 부분이 meta를 지원하도록 수정되어야 합니다.
        # 지금은 이 부분이 동작한다고 가정하고 다음으로 넘어갑니다.
        repositories_output = f"리포지토리 목록 (user: {user_github_url.split('/')[-1]}) - 이 부분은 실제 API 호출로 대체되어야 합니다."

    except Exception as e:
        print(f"Error fetching initial repositories for {user_id}: {e}")
        repositories_output = "Error fetching repositories."

    PROMPT_TEMPLATE = f"""
    You are a portfolio assistant for the user {{user_id}}.
    You are interacting with their GitHub account at {{user_github_url}}.
    Thoroughly answer from the perspective of {{user_id}}.
    Answer questions about their git repositories.
    NEVER answer questions about directories not in a git repository.
    When calling GET tools, use actual existing directories/paths.
    Always base your answers on facts; do not invent information.
    Available repositories:
    {{repositories}}
    """

    prompt_with_user_context = PROMPT_TEMPLATE.format(
        user_id=user_id,
        user_github_url=user_github_url,
        repositories=repositories_output,
    )

    chat_history = deque([], maxlen=5)

    while True:
        command = input(f"[{user_id}] Please enter a command (or 'exit' to quit): ")
        if command.lower() == "exit":
            break

        END_PROMPT_TEMPLATE = f"""
        history: {{chat_history}}
        user: {{command}}"""

        current_prompt = prompt_with_user_context + END_PROMPT_TEMPLATE.format(
            chat_history=list(chat_history), command=command
        )
        print(
            f"\n[DEBUG] Sending prompt to agent for user {user_id}:\n{current_prompt}\n"
        )

        print("\n" + "-" * 40)
        print(f"Running command for {user_id}: {command}")

        # Runner.run 또는 agent.run 호출 시 _meta를 전달할 수 있어야 함
        # agent.set_next_call_meta(meta_for_runner) # 이런 기능이 있다면 매번 설정
        try:
            # result = await Runner.run(starting_agent=agent, input=current_prompt, meta=meta_for_runner) # 가상 meta 파라미터
            # 만약 Runner.run이 meta를 지원하지 않는다면, agent.run이나 agent.call_tool 등을 직접 사용하며 meta를 전달해야 합니다.
            # 지금은 agent가 생성 시점 또는 다른 방식으로 meta(세션 ID)를 MCP 서버에 전달한다고 가정합니다.
            print(
                f"DEBUG: Running agent for {user_id} with session_id {session_id} (implicitly via agent's MCP server config)"
            )
            result = await agent.run(
                input=current_prompt
            )  # agent.run 내부에서 meta가 올바르게 처리되어야 함

            print(f"Output for {user_id}: {result.final_output}")
            current_chat_item = [
                {"role": "user", "content": command},
                {"role": "assistant", "content": result.final_output},
            ]
            chat_history.append(current_chat_item)
        except Exception as e:
            print(f"Error during agent execution for user {user_id}: {e}")
            # 오류 발생 시 사용자에게 알리고 계속 진행할 수 있도록 함
            chat_history.append(
                [
                    {"role": "user", "content": command},
                    {"role": "assistant", "content": f"An error occurred: {e}"},
                ]
            )

        print("\n" + "-" * 40)


async def main_multi_user_simulation():
    # 여러 사용자 시뮬레이션을 위한 정보 (실제 환경에서는 DB 등에서 가져옴)
    # 주의: 실제 PAT를 코드에 하드코딩하지 마세요. 환경 변수나 안전한 저장소에서 로드해야 합니다.
    users_data = {
        "userA": {
            "github_url": "https:#github.com/userA_profile_or_org",
            "pat": os.getenv("USER_A_GITHUB_PAT"),
            "github_api_host": "https:#api.github.com",
        },
        "userB": {
            "github_url": "https:#github.com/userB_profile_or_org",
            "pat": os.getenv("USER_B_GITHUB_PAT"),
            "github_api_host": "https:#custom.ghe.com",
        },  # GHE 사용자 예시
    }
    if not users_data["userA"]["pat"] or not users_data["userB"]["pat"]:
        print(
            "Error: Please set USER_A_GITHUB_PAT and USER_B_GITHUB_PAT environment variables for simulation."
        )
        return

    # MCP 서버 시작 (세션 기반으로 변경됨)
    async with await start_github_mcp_server_session_based() as server:
        print("GitHub MCP Server (Multi-User Session Based) started.")

        # 각 사용자에 대해 포트폴리오 어시스턴트 실행 (비동기적으로 동시에 또는 순차적으로)
        # 여기서는 간단하게 사용자 입력을 번갈아 받는 형태로 시뮬레이션

        # 먼저 모든 사용자의 토큰을 등록합니다.
        for user_id, data in users_data.items():
            print(f"\nRegistering token for {user_id}...")
            client_info = {
                "name": "PortfolioAssistMain",
                "version": "1.0",
            }  # 메인 함수에서 사용할 클라이언트 정보
            await register_user_token(
                server, user_id, data["pat"], data.get("github_api_host"), client_info
            )

        # 사용자 전환 또는 동시 처리를 위한 로직 (여기서는 단순화된 순차적 입력 루프)
        # 실제 서비스에서는 웹 요청 핸들러나 각 사용자의 연결에 따라 컨텍스트를 전환해야 합니다.
        active_user_id = "userA"  # 시작 사용자
        while True:
            print(f"\n--- Current active user: {active_user_id} ---")
            user_data = users_data[active_user_id]

            # run_user_portfolio_assistant 함수는 내부적으로 루프를 돌므로,
            # 여기서는 한 번의 상호작용만 하도록 수정하거나,
            # 사용자가 'switch user' 같은 명령을 입력하면 사용자를 바꾸도록 main_multi_user_simulation 루프를 수정해야 합니다.
            # 아래는 run_user_portfolio_assistant가 단일 명령-응답 사이클을 처리한다고 가정합니다.
            # 또는, 해당 함수를 여기서 직접 호출하지 않고, 각 사용자의 요청이 들어올 때마다 호출하도록 설계 변경.

            # 지금은 단순화를 위해 run_user_portfolio_assistant의 내부 루프를 사용하고,
            # 해당 함수 내에서 'exit' 시 다음 사용자 (또는 프로그램 종료) 로직을 추가하는 것이 나을 수 있습니다.
            # 여기서는 그냥 첫 번째 사용자의 어시스턴트만 실행하고, 'exit'하면 종료되는 형태로 두겠습니다.
            # 실제 다중 사용자 처리는 이 main 함수 레벨에서 사용자 입력을 받아 해당 사용자의 run_user_portfolio_assistant를 호출하는 방식이 될 것입니다.

            await run_user_portfolio_assistant(
                server, active_user_id, user_data["github_url"], user_data["pat"]
            )

            # 사용자 전환 로직 (예시)
            switch_command = input(
                "Enter 'exit' to stop, or type another user ID to switch (e.g., userB): "
            )
            if switch_command.lower() == "exit":
                break
            if switch_command in users_data:
                active_user_id = switch_command
            else:
                print(
                    f"Unknown user ID: {switch_command}. Staying with {active_user_id}."
                )


if __name__ == "__main__":
    if not shutil.which("uvx") and not os.path.exists(
        "./github-mcp-server"
    ):  # 로컬 빌드 파일 체크 추가
        # uvx는 agents 라이브러리 관련 의존성일 수 있으나, 여기서는 Go 서버 실행이 중요.
        # github-mcp-server 실행 파일이 현재 디렉토리에 없다면 에러 발생.
        # 실제 배포 시에는 PATH에 있거나 정확한 경로를 지정해야 합니다.
        raise RuntimeError(
            "github-mcp-server executable not found in current directory. Please build it first."
        )

    asyncio.run(main_multi_user_simulation())
