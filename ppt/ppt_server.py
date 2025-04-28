import asyncio
import shutil

import os
import dotenv
dotenv.load_dotenv(override=True)

from typing import List
from collections import deque

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

print(OPENAI_MODEL)

from agents import Agent, Runner, trace, WebSearchTool
from agents.mcp import MCPServer, MCPServerStdio

import json

with open(r"C:\Users\SSAFY\Desktop\repo\github_mcp\ppt\pptx-compose\example.json", "r", encoding='utf-8') as f:
    template = json.load(f)

async def run(mcp_servers: List[MCPServer]):
    print("Running...")
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about json",
        mcp_servers=mcp_servers,
        tools=[WebSearchTool()]
    )

    print("Agent initialized.")

    PROMPT = f"""
    당신은 json 전문가입니다. 다음 json template는 pptx-compose의 json template입니다.
    이 template을 바탕으로 사용자가 원하는 pptx를 생성하기 위한 json을 생성해 주세요.
    - 항상 template의 구조를 정확하게 유지합니다.
    - template의 내용을 분석하고, 사용자가 원하는 내용을 적절하게 template에 채워 json을 생성합니다.
    사용자가 원하는 pptx의 주제에 맞는 json을 생성해 주세요.
    """

    chat_history = deque([], maxlen=2)

    while True:
        # Ask the user for the git command
        command = input("Please enter a git command (or 'exit' to quit): ")
        if command.lower() == "exit":
            break
        
        END_PROMPT = f"""
        history: {chat_history}
        user: {command}"""
        # prompt = PROMPT.format(chat_history=chat_history) + END_PROMPT + sample
        prompt = PROMPT + END_PROMPT
        print(prompt)

        # Run the command and print the result
        print("\n" + "-" * 40)
        print(f"Running: {command}")
        result = await Runner.run(starting_agent=agent, input=prompt, max_turns=30)
        print(result.final_output)
        current_chat = [
            {"role": "user", "content": command},
            {"role": "assistant", "content": result.final_output},
        ]
        chat_history.append(current_chat)
        print("\n" + "-" * 40)


PPT_MCP_PATH = r"C:\Users\kwon\Desktop\repo\github_mcp\ppt\powerpoint"


def init_servers():
    # Initialize the MCP server with the specified parameters
    print("Initializing PPT servers...")

    server1 = MCPServerStdio(
        # params={
        #     "command": "uv",
        #     "args": ["run", "ppt/main.py"]
        # }
        params={
            "command": "uv",
            # "env": {
            #     "TOGETHER_API_KEY": "api_key"
            # },
            "args": [
                "--directory",
                PPT_MCP_PATH,
                "run",
                "powerpoint",
                "--folder-path",
                "./ppt_result"
            ]
        }
    )

    server1_1 = MCPServerStdio(
        params={
            "command": "npx",
            "args": [
                "-y",
                "@canva/cli@latest",
                "mcp"
            ]
        }
    )

    print("Initializing File System servers...")

    server2 = MCPServerStdio(
        name="Filesystem Server, via npx",
        params={
            "command": "npx",
            "args": ["-y", "@modelcontextprotocol/server-filesystem", "C:/Users/SSAFY/Desktop/repo/github_mcp/ppt"],
        }
    )

    return server1, server1_1, server2

async def main():
    # Ask the user for the directory path
    server1, server1_1, server2 = init_servers()
    print("PPT servers initialized.")
    async with server2 as s2:
        # Create the MCP server and start it
        mcp_servers = [s2]
        with trace(workflow_name="MCP PPT Example"):
            await run(mcp_servers=mcp_servers)

if __name__ == "__main__":
    # if not shutil.which("uv"):
    #     raise RuntimeError("uv is not installed. Please install it with `pip install uv`.")

    asyncio.run(main())
