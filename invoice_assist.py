import os
import pandas as pd
from agno.agent import Agent
from agno.team import Team
from agno.models.google import Gemini
from agno.playground import Playground
import datetime
import gradio as gr
import asyncio
from agno.memory.v2.memory import Memory
from agno.storage.sqlite import SqliteStorage
from agno.memory.v2.db.sqlite import SqliteMemoryDb

from tools import read_timesheet_data, create_invoice_document, save_or_update_timesheet, send_message_with_attachments, get_greeting, set_cell_border

from dotenv import load_dotenv  
load_dotenv(override=True)
google_api_key = os.getenv("GOOGLE_API_KEY")

today = datetime.date.today()

user_id = "user_1"  # static for now; can be dynamic later
db_file = "memeory/invoice_assist.db"  # use persistent location

memory = Memory(
    model=Gemini(id="gemini-2.0-flash", api_key=google_api_key),  # memory summarizer model
    db=SqliteMemoryDb(table_name="user_memories", db_file=db_file),
)

storage = SqliteStorage(table_name="user_info", db_file=db_file)
  
attendence_agent = Agent(
    name="Attendence Assistant",
    model=Gemini(
        id="gemini-2.0-flash",
        api_key=google_api_key
    ),
    tools=[
        read_timesheet_data,
        save_or_update_timesheet
    ],
    instructions = f"""
You manage attendance records in an Excel timesheet.

You can use:
- `read_timesheet_data`: Read the full timesheet or filter by date/week/month.
- `save_or_update_timesheet`: Add or update a row in the timesheet.

Guidelines:
1. The timesheet filename is inferred by the tool based on the current or mentioned month.
    Format for filename: timesheet_<monthName>.xlsx
2. Status options: P = Present, HL = Half Day Leave, L = Leave, WO = Week Off, H = Holiday.
3. Remarks should be ≤ 5 words.
4. For view requests:
   - Use `read_timesheet_data` and present clean summaries (per date/week/month).
5. For add/update requests:
   - First call `read_timesheet_data` to avoid duplicate entries.
   - Then call `save_or_update_timesheet` with the updated row (date, status, remarks).
6. Infer date, month, and intent from user input.

Today's date: {today}.""",
    markdown=True,
)

invoice_agent = Agent(
    name="Invoice Assistant",
    model=Gemini(
        id="gemini-2.5-pro",
        api_key=google_api_key
    ),
    debug_mode=True,
    tools=[
        read_timesheet_data,
        create_invoice_document
    ],
    memory=memory,  # ✅ attach memory here
    enable_agentic_memory=True,  # ✅ allow the agent to extract/update
    enable_user_memories=True,
    add_history_to_messages=False,
    instructions = f"""
You generate monthly invoices based on Excel timesheet data.

Tools:
- `read_timesheet_data`: Fetch attendance data.
- `create_invoice_document`: Create the invoice DOCX file using structured data.

Access the memory to obtain informations such as:
- Full name
- Employee ID
- Department
- Remaining leave balance for the current or past months

Instructions:

1. Use `read_timesheet_data` to get data for the current or user-specified month.
2. Apply business logic:
   - P = Present, 
   - HL = Half Day Leave, 
   - L = Leave, 
   - WO = Week Off, 
   - H = Holiday
   - 2 leaves allowed per month
3. Before generating the invoice:
   - Retrieve the user's name, employee ID, and department from memory.
   - Retrieve the cost of pay per day (in Rs.) from the memory
   - Retrieve the remaining leave balance from memory (or assume 0 if not available).
Note: Computed cost = number of P (Present) days * <pay per full working day> (get pay per full working day from memory)
4. Build the invoice dictionary in the following format:
    {{
        "name": "NAME: <user's full name from memory>",
        "date": "Date: <current date>",
        "bill_to": [
            "PROD SOFTWARE INDIA PRIVATE LIMITED",
            "Kalyani Platina, Ground Floor, Block I, No 24",
            "EPIP Zone Phase II, Whitefield",
            "Bangalore, Karnataka, 560 066"
        ],
        "salary_description": 'Salary for the month of "<month> <year>" payroll',
        "details": [
            "Employee Number: <employee ID from memory>",
            "Department: <department from memory>",
            "Month: <the month of invoice>",
            "Working Days: <calculated working days>",
            "Cumulative Leaves Taken: <L days>",
            "Balance Leaves: <remaining leaves from memory>" //Balance leave should be <remaining leave in the memory + 2 leaves per month>. If any leaves are takes, it should be subrated from this total remaining leaves
        ],
        "total": "<computed total>/-",
        "total_words": "Rs. <computed total in words>"
    }}
5. Save the invoice using `create_invoice_document`:
    - filename: invoice_<month>.docx
    - data: the generated dictionary
6. After computing the invoice:
   - Calculate remaining leaves (2 allowed - number of L days)
   - Update memory with the sentence:
     "My remaining leaves after <month> are <X>."
   - Example: "My remaining leaves after July are 2."

Do not request this information from the user. If it's not found in memory, state clearly that the invoice could not be generated due to missing data.
Infer the current month if not specified.
Today's date: {today}.
""",

    markdown=True,
)

email_agent = Agent(
    name="Email Assistant",
    model=Gemini(
        id="gemini-2.0-flash",
        api_key=google_api_key
    ),
    tools=[
        send_message_with_attachments
    ],
    instructions = f"""
You send messages with timesheet and invoice files attached.

Tool:
- `send_message_with_attachments`: Send a message with XLSX and DOCX files.

Usage:
- The tool auto-generates filenames based on the current or user-specified month:
   - Timesheet: timesheet_<month>.xlsx
   - Invoice: invoice_<month>.docx

Steps:
1. Infer month and recipient (if any).
2. Call `send_message_with_attachments` with:
    {{
        "xlsx_filename": "...",
        "docx_filename": "..."
    }}
3. Confirm success to the user.

Do not request additional input or confirmation.
Today's date: {today}.""",
    markdown=True,
)

team = Team(
    name="Manager",
    members=[attendence_agent, invoice_agent, email_agent],
    mode="route",
    model=Gemini(id="gemini-2.0-flash", api_key=google_api_key),
    instructions="""
    You're a smart team that handles attendance, invoices, and email tasks. Choose the best agent for the user's request. 
    You have permission to reply only if the user asks something from the memory or something about the user has to be added or updated in the memory. 
    You are allowed to remember only important persistent user information such as:
    - Full name
    - Employee ID
    - Department
    - Remaining leave balance

    Do not store or memorize details from general conversations, invoices, or attendance logs. Only extract information if it is clearly a user detail.

    In any other cases, you are don't have the permission to ask for any information, just route the task to other members in the team.
    And you are allowed to reply to generic things like "Hi", "Hello", "What can you do?", etc, any tasks should to passed on to the members no matter what.
    """,
    memory=memory,
    storage=None,
    enable_agentic_memory=True,       # 💾 Store & summarize new memories
    enable_user_memories=True,        # 🧠 Enable retrieval in context
    add_history_to_messages=False,     # 💬 Adds relevant chat history
    num_history_runs=3,
    show_tool_calls=True,
    markdown=True,
    debug_mode=True,
    show_members_responses=True,
)


# # --- Playground App ---
# playground = Playground(agents=[attendence_agent, invoice_agent, email_agent])
# app = playground.get_app()

# # --- Server Entrypoint ---
# if __name__ == "__main__":
#     playground.serve("invoice_assist:app", reload=True)

# # --- CONFIG ---
config = {"user_id": user_id}

# --- CHAT FUNCTION ---
async def chat(user_input, history):
    if not user_input.strip():
        return "Please enter a message."

    # 👇 Only send the latest message to the team (Agno handles internal memory/history)
    result = await team.arun({
        "role": "user",
        "content": user_input
    }, config=config)

    # ✅ Return only the assistant's reply
    return result.content


# --- GRADIO UI ---
gr.ChatInterface(
    fn=chat,
    title="🧾 Invoice Assistant (Agno)",
    description="Ask me to log attendance, generate invoices, or send them via email!",
    theme="soft",
    examples=[
        "What do you know about me?",
        "Mark today as present and that I have worked on AI",
        "Show my attendance for this week",
        "Generate invoice for July",
        "Send the email for July",
    ],
    chatbot=gr.Chatbot(height=450),
    type="messages"  # ✅ Use message-type format
).launch()
