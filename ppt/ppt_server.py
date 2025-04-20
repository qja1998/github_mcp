import asyncio
import shutil

import os
import dotenv
dotenv.load_dotenv()

from collections import deque

GITHUB_PERSONAL_ACCESS_TOKEN = os.getenv("GITHUB_PERSONAL_ACCESS_TOKEN")
OPENAI_MODEL = os.getenv("OPENAI_MODEL")

from agents import Agent, Runner, trace, WebSearchTool, FileSearchTool
from agents.mcp import MCPServer, MCPServerStdio

async def run(mcp_server: MCPServer):
    agent = Agent(
        model=OPENAI_MODEL,
        name="Assistant",
        instructions=f"Answer questions about the notion documents",
        mcp_servers=[mcp_server],
        tools=[WebSearchTool(), FileSearchTool()]
    )

    PROMPT = f"""
    You are a potfolio assistant.
    PPT 형태로 내용들을 정리하여 제공하세요.
    """

    chat_history = deque([], maxlen=5)

    while True:
        # Ask the user for the git command
        command = input("Please enter a git command (or 'exit' to quit): ")
        if command.lower() == "exit":
            break
        
        END_PROMPT = f"""
        history: {chat_history}
        user: {command}"""
        prompt = PROMPT.format(chat_history=chat_history) + END_PROMPT
        print(prompt)

        # Run the command and print the result
        print("\n" + "-" * 40)
        print(f"Running: {command}")
        result = await Runner.run(starting_agent=agent, input=prompt, max_turns=20)
        print(result.final_output)
        current_chat = [
            {"role": "user", "content": command},
            {"role": "assistant", "content": result.final_output},
        ]
        chat_history.append(current_chat)
        print("\n" + "-" * 40)


PPT_MCP_PATH = r"/Users/kwon/Desktop/repository/github_mcp/ppt/Office-PowerPoint-MCP-Server/ppt_mcp_server.py"

async def main():
    # Ask the user for the directory path

    async with MCPServerStdio(
        cache_tools_list=True,  # Cache the tools list, for demonstration
        # params={
        #     "command": "uv",
        #     "env": {
        #         "SD_WEBUI_URL": "http://localhost:7860",
        #         "SD_AUTH_USER": "ppt-mcp",
        #         "SD_AUTH_PASS": "1234",
        #     },
        #     "args": [
        #         "--directory",
        #         PPT_MCP_PATH,
        #         "run",
        #         "powerpoint",
        #         "--folder-path",
        #         "./ppt_result"
        #     ]
        # }
        params={
            "command": "python",
            "args": [PPT_MCP_PATH],
            "env": {}
        }
    ) as server:
        print("setted ppt server")
        with trace(workflow_name="MCP PPT Example"):
            await run(server)


if __name__ == "__main__":
    # if not shutil.which("uv"):
    #     raise RuntimeError("uv is not installed. Please install it with `pip install uv`.")

    asyncio.run(main())
