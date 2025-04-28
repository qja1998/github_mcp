import asyncio
import shutil

import os
import dotenv
from typing import List
dotenv.load_dotenv(override=True)
ROOT = os.path.abspath(os.path.dirname(__file__))

from collections import deque

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

from agents import Agent, Runner, trace
from agents.mcp import MCPServer, MCPServerStdio

async def run(mcp_server: MCPServer, directory_path: str, project_name_list: List[str]):
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about the git repositories, use that for repo_path",
        mcp_servers=[mcp_server],
    )
    
    user_id = directory_path.split("/")[-1]

    repositories = await Runner.run(starting_agent=agent, input=f"{directory_path}의 모든 repository를 ','로 구분해서 나열")

    PROMPT = f"""
    You are a potfolio assistant. You are given {directory_path} (repositories: {repositories}).
    철저하게 {user_id} 관점에서 답변하라.
    NEVER answer questions about the directories that are not in the git repository.
    GET 호출을 할 때는 반드시 실제 있는 디렉토리를 사용해야 한다.
    답변은 절대 지어내지 말고, 항상 사실에 기반해야 한다.
    {project_name_list} 프로젝트의 repository의 모든 내용을 분석하여 다음과 같은 형식으로 정리한다.
    만약 알 수 없는 내용은 format 그대로 사용한다.
    지원자가 포트폴리오를 작성하는 말투로 작성
    commit을 기반으로 사용자가 실제로 기여한 부분만 포함할 것
    **format 내용만을 출력**
    portfolio format:
    """

    FORMAT_PROMPT = """
    ```
    {
        "slogan_main":"recommend slogan for user",  # 프로젝트를 기반으로 유저의 개발자 슬로건을 추천
        "experience_list":["experience", "experience2", ...],  # 프로젝트의 이름 위주로 간단하게 정리
        "award_list":["award1", "award2", ...],  # 수상명 위주로 간단하게 정리
        "vision_slogan":"vision_slogan",  # 프로젝트를 기반으로 유저가 나아갈 비전을 제시
        "vision_description":"vision_description",
        "project_name":"project_name",
        "outline":"outline",  # 30자 이내, 프로젝트의 의미를 간단하게 정리
        "detail_content_list":["What_the_user_did_in_the_project1", "What_the_user_did_in_the_project2", ...],  # 각 항목 20자 이내, 사용자가 한 활동을 어필할 수 있게 정리
        "tech_section":{
            "tech_domain1": ["tech1", "tech2", ...],
            "Backend": ["Django"]  # example
        },
    }
    ```
    answer:
    """

    PROMPT += FORMAT_PROMPT
    result = await Runner.run(starting_agent=agent, input=PROMPT)
    print("Portfolio:", result.final_output)

    chat_history = deque([{"role": "assistant", "content": result.final_output}], maxlen=5)

    while True:
        # Ask the user for the git command
        command = input("Please enter a git command (or 'exit' to quit): ")
        
        if command.lower() == "exit":
            break
        
        END_PROMPT = f"""
        history: {chat_history}
        user: {command}"""

        prompt = PROMPT + END_PROMPT
        print(prompt)

        # Run the command and print the result
        print("\n" + "-" * 40)
        print(f"Running: {command}")
        result = await Runner.run(starting_agent=agent, input=prompt)
        print(result.final_output)
        current_chat = [
            {"role": "user", "content": command},
            {"role": "assistant", "content": result.final_output},
        ]
        chat_history.append(current_chat)
        print("\n" + "-" * 40)


async def main():
    # Ask the user for the directory path
    directory_path = "https://github.com/qja1998"

    async with MCPServerStdio(
        name='Github MCP',
        cache_tools_list=True,  # Cache the tools list, for demonstration
        params={
            "command": "docker",
            "args": [
                "run",
                "-i", "--rm",
                "-e", "GITHUB_PERSONAL_ACCESS_TOKEN",
                "ghcr.io/github/github-mcp-server"
                ],
            "env": {
                "GITHUB_PERSONAL_ACCESS_TOKEN": GITHUB_PERSONAL_ACCESS_TOKEN,
            }
        }
    ) as server:
        print("Run server")
        with trace(workflow_name="MCP Git Example"):
            await run(server, directory_path, "co2-emission-management")


if __name__ == "__main__":
    if not shutil.which("uvx"):
        raise RuntimeError("uvx is not installed. Please install it with `pip install uvx`.")

    asyncio.run(main())