import os
import io
import json
import time
import sqlite3
from datetime import datetime, date, timedelta

import streamlit as st
import pandas as pd
from PIL import Image
from PyPDF2 import PdfReader
from docx import Document

from google import genai
from google.genai import types


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Infinity Bot",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Primary model + fallback models
MODEL_LIST = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
]

DB_FILE = "learning_portal.db"


# ============================================================
# STYLE
# ============================================================

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* ---------- Page ---------- */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background-color: #e0f2fe !important;
    color: #0c4a6e !important;
}

[data-testid="stHeader"],
[data-testid="stBottom"],
[data-testid="stBottom"] > div {
    background-color: #e0f2fe !important;
}

section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div {
    background-color: #bae6fd !important;
}

.block-container {
    max-width: 1250px;
    padding-top: 1.5rem;
}

/* ---------- Text ---------- */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3,
[data-testid="stMarkdownContainer"] h4,
[data-testid="stMarkdownContainer"] h5,
[data-testid="stMarkdownContainer"] h6,
[data-testid="stMarkdownContainer"] strong,
[data-testid="stMarkdownContainer"] em,
[data-testid="stWidgetLabel"] p,
label,
[data-testid="stMetricValue"],
[data-testid="stMetricLabel"],
[data-testid="stMetricLabel"] p {
    color: #0c4a6e !important;
}

[data-testid="stMarkdownContainer"] p code,
[data-testid="stMarkdownContainer"] li code {
    color: #0369a1 !important;
    background-color: #e0f2fe !important;
}

[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] p {
    color: #0369a1 !important;
}

.app-title {
    font-size: 32px;
    font-weight: 700;
    color: #0c4a6e !important;
}

.app-subtitle {
    color: #0369a1 !important;
    margin-bottom: 20px;
}

hr {
    border-color: #7dd3fc !important;
}

/* ---------- Cards ---------- */
.card {
    background: #f0f9ff;
    border: 1px solid #7dd3fc;
    border-radius: 15px;
    padding: 18px;
    margin-bottom: 12px;
}

.course-card {
    background: #f0f9ff;
    border: 1px solid #7dd3fc;
    border-radius: 15px;
    padding: 15px;
    min-height: 145px;
}

/* ---------- Buttons ---------- */
.stButton button,
.stDownloadButton button {
    border-radius: 10px;
    background-color: #f0f9ff !important;
    color: #0369a1 !important;
    border: 1px solid #7dd3fc !important;
}

.stButton button p,
.stDownloadButton button p {
    color: inherit !important;
}

.stButton button:hover,
.stDownloadButton button:hover {
    background-color: #0ea5e9 !important;
    border-color: #0ea5e9 !important;
    color: #ffffff !important;
}

/* ---------- Inputs ---------- */
.stTextInput input,
.stTextArea textarea,
div[data-baseweb="input"],
div[data-baseweb="base-input"],
div[data-baseweb="textarea"],
div[data-baseweb="select"] > div {
    background-color: #f0f9ff !important;
    color: #0c4a6e !important;
    border-color: #7dd3fc !important;
}

div[data-baseweb="select"] * {
    color: #0c4a6e !important;
}

div[data-baseweb="select"] svg {
    fill: #0369a1 !important;
}

div[data-baseweb="popover"] ul,
div[data-baseweb="menu"] {
    background-color: #f0f9ff !important;
}

div[data-baseweb="popover"] li,
div[data-baseweb="popover"] li * {
    color: #0c4a6e !important;
}

div[data-baseweb="popover"] li:hover {
    background-color: #bae6fd !important;
}

/* ---------- Tabs ---------- */
button[data-baseweb="tab"] p {
    color: #0369a1 !important;
}

button[data-baseweb="tab"][aria-selected="true"] p {
    color: #0284c7 !important;
    font-weight: 600;
}

div[data-baseweb="tab-highlight"] {
    background-color: #0ea5e9 !important;
}

div[data-baseweb="tab-border"] {
    background-color: #7dd3fc !important;
}

/* ---------- Progress ---------- */
[data-testid="stProgress"] div[role="progressbar"] {
    background-color: #bae6fd !important;
}

[data-testid="stProgress"] div[role="progressbar"] > div {
    background-color: #0ea5e9 !important;
}

/* ---------- Chat ---------- */
[data-testid="stChatMessage"] {
    background-color: #f0f9ff !important;
    border: 1px solid #bae6fd;
    border-radius: 12px;
}

[data-testid="stChatInput"],
[data-testid="stChatInput"] > div,
[data-testid="stChatInput"] textarea {
    background-color: #f0f9ff !important;
    color: #0c4a6e !important;
    border-color: #7dd3fc !important;
}

/* ---------- Expander / Alerts ---------- */
[data-testid="stExpander"] {
    background-color: #f0f9ff !important;
    border: 1px solid #7dd3fc !important;
    border-radius: 10px;
}

[data-testid="stExpander"] summary p {
    color: #0c4a6e !important;
}

[data-testid="stAlert"],
[data-testid="stAlert"] p {
    color: #0c4a6e !important;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATABASE
# ============================================================

db = sqlite3.connect(
    DB_FILE,
    check_same_thread=False
)

db.row_factory = sqlite3.Row


def init_db():

    cur = db.cursor()

    # USERS
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL
        )
    """)

    # CHAT SESSIONS
    cur.execute("""
        CREATE TABLE IF NOT EXISTS chat_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            title TEXT DEFAULT 'New Chat',
            created_at TEXT,
            updated_at TEXT
        )
    """)

    # CHATS
    cur.execute("""
        CREATE TABLE IF NOT EXISTS chats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            timestamp TEXT,
            session_id INTEGER
        )
    """)

    # SCORES
    cur.execute("""
        CREATE TABLE IF NOT EXISTS lms_scores_v3 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            topic TEXT,
            score_pct REAL,
            correct_count INTEGER,
            total_count INTEGER,
            timestamp TEXT
        )
    """)

    # PROGRESS
    cur.execute("""
        CREATE TABLE IF NOT EXISTS learning_progress (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            language TEXT,
            topic TEXT,
            completed INTEGER DEFAULT 0,
            timestamp TEXT
        )
    """)

    # --------------------------------------------------------
    # CHECK CHAT SESSION COLUMN
    # --------------------------------------------------------

    columns = [
        row["name"]
        for row in cur.execute(
            "PRAGMA table_info(chats)"
        ).fetchall()
    ]

    if "session_id" not in columns:

        cur.execute(
            "ALTER TABLE chats ADD COLUMN session_id INTEGER"
        )

    db.commit()

    # --------------------------------------------------------
    # MIGRATE OLD CHATS
    # --------------------------------------------------------
    # If old chats existed without session_id,
    # attach them to one session per user.
    # Nothing is deleted.
    # --------------------------------------------------------

    users_with_old_chats = cur.execute("""
        SELECT DISTINCT username
        FROM chats
        WHERE session_id IS NULL
    """).fetchall()

    for user_row in users_with_old_chats:

        old_username = user_row["username"]

        existing_session = cur.execute("""
            SELECT id
            FROM chat_sessions
            WHERE username=?
            ORDER BY id ASC
            LIMIT 1
        """, (old_username,)).fetchone()

        if existing_session:

            session_id = existing_session["id"]

        else:

            now = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

            cur.execute("""
                INSERT INTO chat_sessions(
                    username,
                    title,
                    created_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?)
            """, (
                old_username,
                "Previous Chats",
                now,
                now
            ))

            session_id = cur.lastrowid

        cur.execute("""
            UPDATE chats
            SET session_id=?
            WHERE username=?
            AND session_id IS NULL
        """, (
            session_id,
            old_username
        ))

    # Extra feature tables - existing data is preserved
    cur.execute("CREATE TABLE IF NOT EXISTS user_notes (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL, title TEXT NOT NULL, content TEXT NOT NULL, language TEXT, topic TEXT, created_at TEXT, updated_at TEXT)")
    cur.execute("CREATE TABLE IF NOT EXISTS bookmarks (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL, language TEXT NOT NULL, topic TEXT NOT NULL, created_at TEXT, UNIQUE(username, language, topic))")
    cur.execute("CREATE TABLE IF NOT EXISTS activity_log (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL, activity_date TEXT NOT NULL, activity TEXT, UNIQUE(username, activity_date))")
    cur.execute("CREATE TABLE IF NOT EXISTS interview_attempts (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL, language TEXT, level TEXT, score REAL, total INTEGER, timestamp TEXT)")

    db.commit()


init_db()


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "logged_in": False,
    "current_user": None,
    "current_session_id": None,
    "messages": [],
    "quiz": None,
    "quiz_submitted": False,
    "lesson": "",
    "writing_result": "",
    "selected_language": "Python",
    "selected_topic": "Python Basics",
}

for key, value in defaults.items():

    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# GEMINI CLIENT
# ============================================================

def get_client():

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:

        try:
            api_key = st.secrets["GEMINI_API_KEY"]

        except Exception:
            api_key = None

    if not api_key:
        return None

    return genai.Client(
        api_key=api_key
    )


client = get_client()


SYSTEM_PROMPT = """
You are an AI programming tutor and general AI assistant.

You help beginners learn programming.

Programming subjects include:
Python, Java, C, C++, JavaScript, HTML, CSS, SQL,
NumPy, Pandas, Data Analytics, Machine Learning,
Agentic AI, RAG, LangChain, LangGraph, Flask,
FastAPI and Streamlit.

Teaching rules:
- Explain step by step.
- Use simple English.
- Give examples.
- Explain code line by line when useful.
- Give practice tasks.
- Help debug errors.
- Do not assume advanced knowledge.
- If an image or document is attached, analyze it.
- If the learner asks for code, provide complete working code.
- If an error is shown, explain the exact reason and fix it.
"""


# ============================================================
# GEMINI REQUEST WITH RETRY + FALLBACK
# ============================================================

def generate_with_fallback(
    contents,
    system_instruction=None,
    response_mime_type=None
):
    """
    Tries Gemini models one by one.

    3.8 -> 3.7 -> 3.6 -> 3.5

    Temporary 503 / unavailable errors are retried.
    """

    if client is None:

        return None, (
            "Gemini API key not configured.\n\n"
            "Please set GEMINI_API_KEY."
        )

    last_error = None

    for model_name in MODEL_LIST:

        # Try each model twice
        for attempt in range(2):

            try:

                config_kwargs = {}

                if system_instruction:
                    config_kwargs["system_instruction"] = (
                        system_instruction
                    )

                if response_mime_type:
                    config_kwargs["response_mime_type"] = (
                        response_mime_type
                    )

                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        **config_kwargs
                    )
                )

                text = (
                    response.text
                    if response
                    else None
                )

                if text:

                    return text, None

                last_error = (
                    f"{model_name}: Empty response"
                )

            except Exception as e:

                error_text = str(e)
                last_error = error_text

                # ------------------------------------------------
                # RETRY TEMPORARY SERVER ERRORS
                # ------------------------------------------------

                temporary_errors = [
                    "503",
                    "UNAVAILABLE",
                    "429",
                    "RESOURCE_EXHAUSTED",
                    "500",
                    "INTERNAL"
                ]

                is_temporary = any(
                    item in error_text.upper()
                    for item in temporary_errors
                )

                if is_temporary:

                    if attempt == 0:

                        time.sleep(2)

                        continue

                    # Move to next model
                    break

                # ------------------------------------------------
                # For non-temporary errors,
                # try next model anyway.
                # ------------------------------------------------

                break

    return None, (
        "Gemini is temporarily unavailable.\n\n"
        "I tried these models:\n"
        + "\n".join(
            f"- {model}"
            for model in MODEL_LIST
        )
        + "\n\nLast error:\n"
        + str(last_error)
    )


# ============================================================
# AUTH FUNCTIONS
# ============================================================

def register_user(username, password):

    username = username.strip()

    if not username or not password:
        return False, "Username and password are required."

    try:

        db.execute(
            """
            INSERT INTO users(username, password)
            VALUES (?, ?)
            """,
            (
                username,
                password
            )
        )

        db.commit()

        return True, "Account created successfully."

    except sqlite3.IntegrityError:

        return False, "Username already exists."


def check_login(username, password):

    row = db.execute(
        """
        SELECT username
        FROM users
        WHERE username=?
        AND password=?
        """,
        (
            username.strip(),
            password
        )
    ).fetchone()

    return row is not None


# ============================================================
# CHAT SESSION FUNCTIONS
# ============================================================

def create_session(username):

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    cur = db.execute(
        """
        INSERT INTO chat_sessions(
            username,
            title,
            created_at,
            updated_at
        )
        VALUES (?, ?, ?, ?)
        """,
        (
            username,
            "New Chat",
            now,
            now
        )
    )

    db.commit()

    return cur.lastrowid


def get_sessions(username):

    return db.execute(
        """
        SELECT *
        FROM chat_sessions
        WHERE username=?
        ORDER BY updated_at DESC, id DESC
        """,
        (
            username,
        )
    ).fetchall()


def load_messages(
    username,
    session_id
):

    rows = db.execute(
        """
        SELECT id,
               role,
               content,
               timestamp
        FROM chats
        WHERE username=?
        AND session_id=?
        ORDER BY id ASC
        """,
        (
            username,
            session_id
        )
    ).fetchall()

    return [
        dict(row)
        for row in rows
    ]


def save_message(
    username,
    session_id,
    role,
    content
):

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    db.execute(
        """
        INSERT INTO chats(
            username,
            role,
            content,
            timestamp,
            session_id
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            username,
            role,
            content,
            now,
            session_id
        )
    )

    db.execute(
        """
        UPDATE chat_sessions
        SET updated_at=?
        WHERE id=?
        AND username=?
        """,
        (
            now,
            session_id,
            username
        )
    )

    db.commit()


def delete_message(
    message_id,
    username
):

    db.execute(
        """
        DELETE FROM chats
        WHERE id=?
        AND username=?
        """,
        (
            message_id,
            username
        )
    )

    db.commit()


def delete_session(
    session_id,
    username
):

    db.execute(
        """
        DELETE FROM chats
        WHERE session_id=?
        AND username=?
        """,
        (
            session_id,
            username
        )
    )

    db.execute(
        """
        DELETE FROM chat_sessions
        WHERE id=?
        AND username=?
        """,
        (
            session_id,
            username
        )
    )

    db.commit()


def clear_all_chats(username):

    db.execute(
        "DELETE FROM chats WHERE username=?",
        (
            username,
        )
    )

    db.execute(
        "DELETE FROM chat_sessions WHERE username=?",
        (
            username,
        )
    )

    db.commit()


def chat_title(text):

    text = " ".join(
        text.strip().split()
    )

    if not text:
        return "New Chat"

    return text[:45]


def rename_session(
    session_id,
    username,
    title
):

    db.execute(
        """
        UPDATE chat_sessions
        SET title=?
        WHERE id=?
        AND username=?
        """,
        (
            title,
            session_id,
            username
        )
    )

    db.commit()


# ============================================================
# PROGRAMMING COURSES
# ============================================================

COURSES = {

    "Python": {
        "icon": "🐍",
        "topics": [
            "Python Basics",
            "Variables & Data Types",
            "Input & Output",
            "Conditions",
            "Loops",
            "Functions",
            "Lists & Tuples",
            "Dictionaries & Sets",
            "Strings",
            "Exception Handling",
            "File Handling",
            "OOP",
            "Modules & Packages",
            "NumPy",
            "Pandas",
            "Mini Project"
        ]
    },

    "Java": {
        "icon": "☕",
        "topics": [
            "Java Basics",
            "Variables & Data Types",
            "Input & Output",
            "Conditions",
            "Loops",
            "Arrays",
            "Methods",
            "Strings",
            "Classes & Objects",
            "Constructors",
            "Inheritance",
            "Polymorphism",
            "Encapsulation",
            "Abstraction",
            "Exception Handling",
            "Mini Project"
        ]
    },

    "C": {
        "icon": "©️",
        "topics": [
            "C Basics",
            "Variables & Data Types",
            "Input & Output",
            "Operators",
            "Conditions",
            "Loops",
            "Arrays",
            "Strings",
            "Functions",
            "Pointers",
            "Structures",
            "File Handling",
            "Mini Project"
        ]
    },

    "C++": {
        "icon": "⚙️",
        "topics": [
            "C++ Basics",
            "Variables & Data Types",
            "Input & Output",
            "Conditions",
            "Loops",
            "Arrays",
            "Functions",
            "Strings",
            "Classes & Objects",
            "Constructors",
            "Inheritance",
            "Polymorphism",
            "STL",
            "Mini Project"
        ]
    },

    "JavaScript": {
        "icon": "🟨",
        "topics": [
            "JavaScript Basics",
            "Variables",
            "Data Types",
            "Operators",
            "Conditions",
            "Loops",
            "Functions",
            "Arrays",
            "Objects",
            "DOM",
            "Events",
            "Async JavaScript",
            "Fetch API",
            "Mini Project"
        ]
    },

    "HTML": {
        "icon": "🌐",
        "topics": [
            "HTML Basics",
            "Page Structure",
            "Headings & Paragraphs",
            "Links & Images",
            "Lists",
            "Tables",
            "Forms",
            "Semantic HTML",
            "Accessibility",
            "Mini Project"
        ]
    },

    "CSS": {
        "icon": "🎨",
        "topics": [
            "CSS Basics",
            "Selectors",
            "Colors & Fonts",
            "Box Model",
            "Display",
            "Flexbox",
            "Grid",
            "Positioning",
            "Responsive Design",
            "Animations",
            "Mini Project"
        ]
    },

    "SQL": {
        "icon": "🗄️",
        "topics": [
            "SQL Basics",
            "SELECT",
            "WHERE",
            "ORDER BY",
            "GROUP BY",
            "HAVING",
            "Aggregate Functions",
            "INSERT",
            "UPDATE",
            "DELETE",
            "JOINS",
            "Subqueries",
            "Window Functions",
            "Mini Project"
        ]
    }
}


# ============================================================
# PROGRESS
# ============================================================

def mark_completed(
    username,
    language,
    topic
):

    existing = db.execute(
        """
        SELECT id
        FROM learning_progress
        WHERE username=?
        AND language=?
        AND topic=?
        """,
        (
            username,
            language,
            topic
        )
    ).fetchone()

    now = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    if existing:

        db.execute(
            """
            UPDATE learning_progress
            SET completed=1,
                timestamp=?
            WHERE id=?
            """,
            (
                now,
                existing["id"]
            )
        )

    else:

        db.execute(
            """
            INSERT INTO learning_progress(
                username,
                language,
                topic,
                completed,
                timestamp
            )
            VALUES (?, ?, ?, 1, ?)
            """,
            (
                username,
                language,
                topic,
                now
            )
        )

    db.commit()


def completed_topics(
    username,
    language
):

    rows = db.execute(
        """
        SELECT topic
        FROM learning_progress
        WHERE username=?
        AND language=?
        AND completed=1
        """,
        (
            username,
            language
        )
    ).fetchall()

    return {
        row["topic"]
        for row in rows
    }


# ============================================================
# FILE READER
# ============================================================

IMAGE_EXTENSIONS = {
    "png",
    "jpg",
    "jpeg",
    "webp"
}

TEXT_EXTENSIONS = {
    "txt",
    "md",
    "py",
    "java",
    "c",
    "cpp",
    "h",
    "hpp",
    "js",
    "ts",
    "html",
    "css",
    "sql",
    "json",
    "xml",
    "yaml",
    "yml"
}


def read_pdf(data):

    reader = PdfReader(
        io.BytesIO(data)
    )

    result = []

    for page in reader.pages:

        result.append(
            page.extract_text() or ""
        )

    return "\n".join(result)


def read_docx(data):

    doc = Document(
        io.BytesIO(data)
    )

    result = []

    for paragraph in doc.paragraphs:

        if paragraph.text.strip():

            result.append(
                paragraph.text
            )

    return "\n".join(result)


def process_files(files):

    text_parts = []
    image_parts = []
    names = []

    for file in files:

        name = file.name

        names.append(name)

        extension = (
            name.lower().rsplit(
                ".",
                1
            )[-1]
            if "." in name
            else ""
        )

        data = file.getvalue()

        # ----------------------------------------------------
        # IMAGE
        # ----------------------------------------------------

        if extension in IMAGE_EXTENSIONS:

            mime = {
                "png": "image/png",
                "jpg": "image/jpeg",
                "jpeg": "image/jpeg",
                "webp": "image/webp"
            }[extension]

            image_parts.append(
                types.Part.from_bytes(
                    data=data,
                    mime_type=mime
                )
            )

        # ----------------------------------------------------
        # PDF
        # ----------------------------------------------------

        elif extension == "pdf":

            text_parts.append(
                f"\n===== {name} =====\n"
                + read_pdf(data)[:50000]
            )

        # ----------------------------------------------------
        # DOCX
        # ----------------------------------------------------

        elif extension == "docx":

            text_parts.append(
                f"\n===== {name} =====\n"
                + read_docx(data)[:50000]
            )

        # ----------------------------------------------------
        # CSV
        # ----------------------------------------------------

        elif extension == "csv":

            try:

                df = pd.read_csv(
                    io.BytesIO(data)
                )

                text_parts.append(
                    f"\n===== {name} =====\n"
                    + df.to_csv(
                        index=False
                    )[:50000]
                )

            except Exception:

                text_parts.append(
                    data.decode(
                        "utf-8",
                        errors="ignore"
                    )[:50000]
                )

        # ----------------------------------------------------
        # TEXT / CODE
        # ----------------------------------------------------

        elif extension in TEXT_EXTENSIONS:

            text_parts.append(
                f"\n===== {name} =====\n"
                + data.decode(
                    "utf-8",
                    errors="ignore"
                )[:50000]
            )

        else:

            text_parts.append(
                f"\n[Unsupported file: {name}]"
            )

    return (
        "\n".join(text_parts),
        image_parts,
        names
    )


# ============================================================
# GEMINI CHAT
# ============================================================

def ask_ai(
    previous_messages,
    user_text,
    file_text="",
    image_parts=None,
    instruction=""
):

    if client is None:

        return (
            "⚠️ Gemini API key not configured.\n\n"
            "Please set GEMINI_API_KEY."
        )

    contents = []

    # --------------------------------------------------------
    # PREVIOUS CHAT
    # --------------------------------------------------------

    for message in previous_messages[-25:]:

        role = (
            "user"
            if message["role"] == "user"
            else "model"
        )

        contents.append(
            types.Content(
                role=role,
                parts=[
                    types.Part.from_text(
                        text=message["content"]
                    )
                ]
            )
        )

    # --------------------------------------------------------
    # CURRENT MESSAGE
    # --------------------------------------------------------

    parts = []

    if user_text:

        parts.append(
            types.Part.from_text(
                text=user_text
            )
        )

    if file_text:

        parts.append(
            types.Part.from_text(
                text=(
                    "\n\nATTACHED FILE CONTENT:\n"
                    + file_text
                )
            )
        )

    if image_parts:

        parts.extend(
            image_parts
        )

    if not parts:

        parts.append(
            types.Part.from_text(
                text="Hello"
            )
        )

    contents.append(
        types.Content(
            role="user",
            parts=parts
        )
    )

    system = SYSTEM_PROMPT

    if instruction:

        system += (
            "\n\nSPECIAL INSTRUCTION:\n"
            + instruction
        )

    answer, error = generate_with_fallback(
        contents=contents,
        system_instruction=system
    )

    if error:

        return (
            "❌ Gemini error:\n\n"
            + error
        )

    return answer


# ============================================================
# QUIZ
# ============================================================

def create_quiz(
    language,
    topic
):

    if client is None:

        return (
            None,
            "Gemini API key not configured."
        )

    prompt = f"""
Create exactly 5 MCQ questions.

Programming Language:
{language}

Topic:
{topic}

Rules:
- Beginner-friendly
- No duplicate questions
- Four options per question
- Correct answer must exactly match one option

Return ONLY JSON.

Format:

{{
    "questions": [
        {{
            "question": "Question",
            "options": [
                "A",
                "B",
                "C",
                "D"
            ],
            "answer": "A",
            "explanation": "Short explanation"
        }}
    ]
}}
"""

    answer, error = generate_with_fallback(
        contents=prompt,
        response_mime_type="application/json"
    )

    if error:

        return None, error

    try:

        quiz = json.loads(
            answer
        )

        return quiz, None

    except Exception as e:

        return (
            None,
            "Quiz JSON parsing error: "
            + str(e)
            + "\n\nGemini response:\n"
            + str(answer)
        )


def save_score(
    username,
    topic,
    correct,
    total
):

    percentage = (
        correct / total * 100
        if total
        else 0
    )

    db.execute(
        """
        INSERT INTO lms_scores_v3(
            username,
            topic,
            score_pct,
            correct_count,
            total_count,
            timestamp
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            username,
            topic,
            percentage,
            correct,
            total,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )
    )

    db.commit()


# ============================================================
# EXTRA FEATURE HELPERS
# ============================================================

def log_activity(username, activity="app"):
    db.execute("INSERT OR IGNORE INTO activity_log(username, activity_date, activity) VALUES (?, ?, ?)",(username,date.today().isoformat(),activity)); db.commit()

def get_streak(username):
    rows=db.execute("SELECT activity_date FROM activity_log WHERE username=?",(username,)).fetchall(); days={r["activity_date"] for r in rows}; d=date.today(); n=0
    while d.isoformat() in days: n+=1; d-=timedelta(days=1)
    return n

def get_quiz_stats(username):
    return dict(db.execute("SELECT COUNT(*) attempts, COALESCE(AVG(score_pct),0) avg_score, COALESCE(MAX(score_pct),0) best_score FROM lms_scores_v3 WHERE username=?",(username,)).fetchone())

def get_total_completed(username):
    return int(db.execute("SELECT COUNT(*) c FROM learning_progress WHERE username=? AND completed=1",(username,)).fetchone()["c"] or 0)

def is_bookmarked(username,language,topic):
    return db.execute("SELECT 1 FROM bookmarks WHERE username=? AND language=? AND topic=?",(username,language,topic)).fetchone() is not None

def toggle_bookmark(username,language,topic):
    row=db.execute("SELECT id FROM bookmarks WHERE username=? AND language=? AND topic=?",(username,language,topic)).fetchone()
    if row: db.execute("DELETE FROM bookmarks WHERE id=?",(row["id"],)); db.commit(); return False
    db.execute("INSERT OR IGNORE INTO bookmarks(username,language,topic,created_at) VALUES(?,?,?,?)",(username,language,topic,datetime.now().strftime("%Y-%m-%d %H:%M:%S"))); db.commit(); return True

def save_note(username,title,content,language,topic):
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S"); db.execute("INSERT INTO user_notes(username,title,content,language,topic,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(username,title,content,language,topic,now,now)); db.commit()

def load_notes(username): return db.execute("SELECT * FROM user_notes WHERE username=? ORDER BY id DESC",(username,)).fetchall()

def delete_note(note_id,username): db.execute("DELETE FROM user_notes WHERE id=? AND username=?",(note_id,username)); db.commit()

def safe_json(text):
    text=(text or "").strip().replace("```json","").replace("```","").strip()
    try: return json.loads(text)
    except Exception:
        for a,b in [("[","]"),("{","}")]:
            i,j=text.find(a),text.rfind(b)
            if i>=0 and j>i: return json.loads(text[i:j+1])
        raise ValueError("AI returned invalid JSON")


# ============================================================
# LOGIN SCREEN
# ============================================================

def login_screen():

    st.markdown(
        '<div class="app-title">🤖 Infinity Bot</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="app-subtitle">
        Programming Learning • AI Tutor •
        Grammar Checker • File Analysis • Practice
        </div>
        """,
        unsafe_allow_html=True
    )

    left, center, right = st.columns(
        [1, 2, 1]
    )

    with center:

        login, register = st.tabs(
            [
                "🔐 Login",
                "📝 Register"
            ]
        )

        # ----------------------------------------------------
        # LOGIN
        # ----------------------------------------------------

        with login:

            username = st.text_input(
                "Username",
                key="login_user"
            )

            password = st.text_input(
                "Password",
                type="password",
                key="login_pass"
            )

            if st.button(
                "Login",
                use_container_width=True,
                key="login_btn"
            ):

                if check_login(
                    username,
                    password
                ):

                    st.session_state.logged_in = True

                    st.session_state.current_user = (
                        username.strip()
                    )

                    sessions = get_sessions(
                        st.session_state.current_user
                    )

                    if sessions:

                        sid = sessions[0]["id"]

                    else:

                        sid = create_session(
                            st.session_state.current_user
                        )

                    st.session_state.current_session_id = sid

                    st.session_state.messages = (
                        load_messages(
                            st.session_state.current_user,
                            sid
                        )
                    )

                    st.rerun()

                else:

                    st.error(
                        "Invalid username or password."
                    )

        # ----------------------------------------------------
        # REGISTER
        # ----------------------------------------------------

        with register:

            new_user = st.text_input(
                "Username",
                key="register_user"
            )

            new_pass = st.text_input(
                "Password",
                type="password",
                key="register_pass"
            )

            confirm = st.text_input(
                "Confirm Password",
                type="password",
                key="register_confirm"
            )

            if st.button(
                "Create Account",
                use_container_width=True,
                key="register_btn"
            ):

                if new_pass != confirm:

                    st.error(
                        "Passwords do not match."
                    )

                else:

                    ok, msg = register_user(
                        new_user,
                        new_pass
                    )

                    if ok:

                        st.success(
                            msg
                        )

                    else:

                        st.error(
                            msg
                        )


# ============================================================
# LOGIN CHECK
# ============================================================

if not st.session_state.logged_in:

    login_screen()

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

username = st.session_state.current_user
log_activity(username, "login")

with st.sidebar:

    st.markdown(
        "## 🤖 Infinity Bot"
    )

    st.caption(
        f"👤 {username}"
    )

    st.divider()

    st.markdown(
        "### 📚 Programming"
    )

    languages = list(
        COURSES.keys()
    )

    language = st.selectbox(
        "Choose Language",
        languages,
        index=(
            languages.index(
                st.session_state.selected_language
            )
            if st.session_state.selected_language
            in languages
            else 0
        )
    )

    st.session_state.selected_language = language

    topics = COURSES[
        language
    ]["topics"]

    topic = st.selectbox(
        "Choose Topic",
        topics,
        index=(
            topics.index(
                st.session_state.selected_topic
            )
            if st.session_state.selected_topic
            in topics
            else 0
        )
    )

    st.session_state.selected_topic = topic

    completed = completed_topics(
        username,
        language
    )

    progress = (
        len(completed) / len(topics)
        if topics
        else 0
    )

    st.progress(
        progress
    )

    st.caption(
        f"{len(completed)} / {len(topics)} completed"
    )

    st.divider()

    # --------------------------------------------------------
    # NEW CHAT
    # --------------------------------------------------------

    if st.button(
        "➕ New Chat",
        use_container_width=True,
        key="new_chat"
    ):

        sid = create_session(
            username
        )

        st.session_state.current_session_id = sid
        st.session_state.messages = []

        st.rerun()

    # --------------------------------------------------------
    # HISTORY
    # --------------------------------------------------------

    st.markdown(
        "### 💬 History"
    )

    sessions = get_sessions(
        username
    )

    for session in sessions:

        sid = session["id"]

        c1, c2 = st.columns(
            [5, 1]
        )

        with c1:

            if st.button(
                session["title"],
                key=f"open_{sid}",
                use_container_width=True
            ):

                st.session_state.current_session_id = sid

                st.session_state.messages = (
                    load_messages(
                        username,
                        sid
                    )
                )

                st.rerun()

        with c2:

            if st.button(
                "🗑️",
                key=f"delete_{sid}"
            ):

                delete_session(
                    sid,
                    username
                )

                remaining = get_sessions(
                    username
                )

                if remaining:

                    new_sid = remaining[0]["id"]

                else:

                    new_sid = create_session(
                        username
                    )

                st.session_state.current_session_id = new_sid

                st.session_state.messages = (
                    load_messages(
                        username,
                        new_sid
                    )
                )

                st.rerun()

    st.divider()

    # --------------------------------------------------------
    # CLEAR ALL
    # --------------------------------------------------------

    if st.button(
        "🧹 Clear All History",
        use_container_width=True,
        key="clear_all"
    ):

        clear_all_chats(
            username
        )

        sid = create_session(
            username
        )

        st.session_state.current_session_id = sid
        st.session_state.messages = []

        st.rerun()

    st.divider()

    # --------------------------------------------------------
    # LOGOUT
    # --------------------------------------------------------

    if st.button(
        "🚪 Logout",
        use_container_width=True,
        key="logout"
    ):

        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.session_state.current_session_id = None
        st.session_state.messages = []

        st.rerun()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="app-title">🤖 Infinity Bot</div>',
    unsafe_allow_html=True
)

st.markdown(
    f"""
    <div class="app-subtitle">
    {COURSES[language]["icon"]}
    {language} → {topic}
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# TABS
# ============================================================

(
    home_tab, learn_tab, practice_tab, chat_tab, writing_tab, progress_tab,
    roadmap_tab, codelab_tab, interview_tab, notes_tab, certificate_tab
) = st.tabs([
    "🏠 Home", "📚 Learn", "🧪 Practice", "💬 AI Chat", "✍️ Grammar", "📊 Progress",
    "🗺 Roadmap", "🧑‍💻 Code Lab", "🎤 Interview", "📝 Notes", "🏆 Certificate"
])


# ============================================================
# HOME
# ============================================================

with home_tab:

    st.subheader(
        "👋 Welcome!"
    )

    stats=get_quiz_stats(username)
    m1,m2,m3,m4,m5=st.columns(5)
    m1.metric("Topics Done",get_total_completed(username))
    m2.metric("Quiz Attempts",int(stats["attempts"]))
    m3.metric("Quiz Average",f"{stats['avg_score']:.0f}%")
    m4.metric("Best Quiz",f"{stats['best_score']:.0f}%")
    m5.metric("🔥 Streak",f"{get_streak(username)} days")
    st.info(f"▶️ Continue with **{language} → {topic}**. Open the Learn tab to continue.")

    st.write(
        "Learn programming step by step with your AI tutor."
    )

    cols = st.columns(4)

    features = [
        (
            "📚",
            "Learn",
            "Programming lessons"
        ),
        (
            "🤖",
            "AI Tutor",
            "Ask coding doubts"
        ),
        (
            "🧪",
            "Practice",
            "AI quizzes"
        ),
        (
            "✍️",
            "Grammar",
            "Improve your English"
        ),
    ]

    for col, item in zip(
        cols,
        features
    ):

        with col:

            st.markdown(
                f"""
                <div class="card">
                    <h2>{item[0]}</h2>
                    <h4>{item[1]}</h4>
                    <p>{item[2]}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.subheader(
        "💻 Programming Languages"
    )

    cols = st.columns(4)

    for index, (
        name,
        data
    ) in enumerate(
        COURSES.items()
    ):

        with cols[index % 4]:

            st.markdown(
                f"""
                <div class="course-card">
                    <h2>
                    {data["icon"]} {name}
                    </h2>
                    <p>
                    {len(data["topics"])} topics
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )


# ============================================================
# LEARN
# ============================================================

with learn_tab:

    st.subheader(
        f"📚 {language} — {topic}"
    )

    if is_bookmarked(username,language,topic):
        if st.button("🔖 Remove Bookmark",key="remove_bookmark"): toggle_bookmark(username,language,topic); st.rerun()
    else:
        if st.button("🔖 Bookmark Topic",key="add_bookmark"): toggle_bookmark(username,language,topic); st.rerun()

    st.write(
        f"Learn **{topic}** from beginner level."
    )

    if st.button(
        "📖 Generate Lesson",
        use_container_width=True,
        key="generate_lesson"
    ):

        with st.spinner(
            "Preparing lesson..."
        ):

            st.session_state.lesson = ask_ai(
                [],
                f"Teach me {topic}.",
                instruction=f"""
Teach {language} programming topic:
{topic}

Include:

1. What it is
2. Why it is used
3. Syntax
4. Simple example
5. Line-by-line explanation
6. Common mistakes
7. Interview question
8. Practice task
"""
            )

    if st.session_state.lesson:

        st.divider()

        st.markdown(
            st.session_state.lesson
        )

        if st.button(
            "✅ Mark Topic Completed",
            key="complete_topic"
        ):

            mark_completed(
                username,
                language,
                topic
            )

            st.success(
                "Topic completed!"
            )


# ============================================================
# PRACTICE
# ============================================================

with practice_tab:

    st.subheader(
        f"🧪 {language} Practice"
    )

    st.info(
        f"Current topic: {topic}"
    )

    if st.button(
        "✨ Generate Quiz",
        use_container_width=True,
        key="generate_quiz"
    ):

        with st.spinner(
            "Generating questions..."
        ):

            quiz, error = create_quiz(
                language,
                topic
            )

        if error:

            st.error(
                error
            )

        else:

            st.session_state.quiz = quiz
            st.session_state.quiz_submitted = False

            st.rerun()

    quiz = st.session_state.quiz

    if quiz:

        answers = {}

        questions = quiz.get(
            "questions",
            []
        )

        for i, question in enumerate(
            questions
        ):

            st.markdown(
                f"""
                ### {i + 1}. {question["question"]}
                """
            )

            answers[i] = st.radio(
                "Choose answer",
                question["options"],
                index=None,
                key=f"quiz_{i}"
            )

        if st.button(
            "✅ Submit Quiz",
            use_container_width=True,
            disabled=st.session_state.quiz_submitted,
            key="submit_quiz"
        ):

            correct = 0

            for i, question in enumerate(
                questions
            ):

                if (
                    answers.get(i)
                    == question["answer"]
                ):

                    correct += 1

            total = len(
                questions
            )

            percentage = (
                correct / total * 100
                if total
                else 0
            )

            save_score(
                username,
                f"{language} - {topic}",
                correct,
                total
            )

            st.session_state.quiz_submitted = True

            st.success(
                f"🎉 Score: {correct}/{total} "
                f"({percentage:.0f}%)"
            )

            for i, question in enumerate(
                questions
            ):

                if (
                    answers.get(i)
                    == question["answer"]
                ):

                    st.success(
                        f"Q{i + 1}: Correct"
                    )

                else:

                    st.error(
                        f"Q{i + 1}: Correct answer = "
                        f"{question['answer']}"
                    )

                st.caption(
                    question.get(
                        "explanation",
                        ""
                    )
                )


# ============================================================
# AI CHAT
# ============================================================

with chat_tab:

    st.subheader(
        "💬 AI Tutor"
    )

    st.caption(
        "Ask anything • Upload multiple files • "
        "Images • Code • PDF • DOCX • CSV"
    )

    # --------------------------------------------------------
    # EXISTING MESSAGES
    # --------------------------------------------------------

    for message in st.session_state.messages:

        role = (
            "assistant"
            if message["role"] == "assistant"
            else "user"
        )

        with st.chat_message(
            role
        ):

            st.markdown(
                message["content"]
            )

            if role == "user":

                if st.button(
                    "🗑️ Delete",
                    key=f"msg_delete_{message['id']}"
                ):

                    delete_message(
                        message["id"],
                        username
                    )

                    st.session_state.messages = (
                        load_messages(
                            username,
                            st.session_state.current_session_id
                        )
                    )

                    st.rerun()

    # --------------------------------------------------------
    # CHAT INPUT WITH ATTACHMENT
    # --------------------------------------------------------

    submission = st.chat_input(
        "Message your AI tutor...",
        accept_file="multiple",
        file_type=[
            "pdf",
            "docx",
            "txt",
            "md",
            "csv",
            "py",
            "java",
            "c",
            "cpp",
            "h",
            "hpp",
            "js",
            "ts",
            "html",
            "css",
            "sql",
            "json",
            "xml",
            "yaml",
            "yml",
            "png",
            "jpg",
            "jpeg",
            "webp"
        ],
        max_upload_size=200,
        key="main_chat"
    )

    if submission:

        user_text = (
            getattr(
                submission,
                "text",
                ""
            )
            or ""
        ).strip()

        files = (
            getattr(
                submission,
                "files",
                []
            )
            or []
        )

        if not user_text and not files:

            st.stop()

        file_text, image_parts, names = (
            process_files(
                files
            )
        )

        display_message = (
            user_text
            if user_text
            else "📎 Uploaded files"
        )

        if names:

            display_message += (
                "\n\n📎 **Attached:** "
                + ", ".join(names)
            )

        # Save current history BEFORE AI call
        old_messages = (
            st.session_state.messages.copy()
        )

        # ----------------------------------------------------
        # SAVE USER MESSAGE
        # ----------------------------------------------------

        log_activity(username, "chat")
        save_message(
            username,
            st.session_state.current_session_id,
            "user",
            display_message
        )

        # ----------------------------------------------------
        # AUTO TITLE
        # ----------------------------------------------------

        row = db.execute(
            """
            SELECT title
            FROM chat_sessions
            WHERE id=?
            AND username=?
            """,
            (
                st.session_state.current_session_id,
                username
            )
        ).fetchone()

        if (
            row
            and row["title"] == "New Chat"
        ):

            rename_session(
                st.session_state.current_session_id,
                username,
                chat_title(
                    user_text
                    or (
                        names[0]
                        if names
                        else "New Chat"
                    )
                )
            )

        # ----------------------------------------------------
        # DISPLAY USER
        # ----------------------------------------------------

        with st.chat_message(
            "user"
        ):

            st.markdown(
                display_message
            )

            for file in files:

                extension = (
                    file.name.lower()
                    .rsplit(
                        ".",
                        1
                    )[-1]
                )

                if extension in IMAGE_EXTENSIONS:

                    try:

                        st.image(
                            Image.open(
                                io.BytesIO(
                                    file.getvalue()
                                )
                            ),
                            caption=file.name,
                            width=300
                        )

                    except Exception:

                        pass

        # ----------------------------------------------------
        # AI RESPONSE
        # ----------------------------------------------------

        with st.chat_message(
            "assistant"
        ):

            with st.spinner(
                "Thinking..."
            ):

                answer = ask_ai(
                    old_messages,
                    user_text,
                    file_text,
                    image_parts,
                    instruction=f"""
The learner is currently studying:

{language} → {topic}

Answer according to the learner's current level.
"""
                )

            st.markdown(
                answer
            )

        # ----------------------------------------------------
        # SAVE AI RESPONSE
        # ----------------------------------------------------

        save_message(
            username,
            st.session_state.current_session_id,
            "assistant",
            answer
        )

        st.session_state.messages = (
            load_messages(
                username,
                st.session_state.current_session_id
            )
        )

        st.rerun()


# ============================================================
# GRAMMAR / WRITING
# ============================================================

with writing_tab:

    st.subheader(
        "✍️ Grammar & Writing Assistant"
    )

    action = st.selectbox(
        "What do you want to do?",
        [
            "Grammar Checker",
            "Spelling Checker",
            "Rewrite",
            "Professional",
            "Friendly",
            "Simple English",
            "Shorten",
            "Expand",
            "Summarize",
            "Improve Writing",
            "Translate"
        ],
        key="writing_action"
    )

    target = "English"

    if action == "Translate":

        target = st.selectbox(
            "Translate to",
            [
                "English",
                "Tamil",
                "Hindi",
                "Malayalam",
                "Telugu",
                "Kannada",
                "French",
                "German",
                "Spanish",
                "Japanese"
            ],
            key="translate_target"
        )

    text = st.text_area(
        "Enter your text",
        height=220,
        placeholder=(
            "Type or paste your text here..."
        ),
        key="writing_text"
    )

    if st.button(
        "✨ Check / Improve",
        use_container_width=True,
        key="writing_button"
    ):

        if not text.strip():

            st.warning(
                "Please enter some text."
            )

        else:

            instruction_map = {

                "Grammar Checker":
                    """
Check grammar, spelling, punctuation and sentence
structure.

Return:
1. Corrected text
2. Important corrections
3. Short explanation
""",

                "Spelling Checker":
                    """
Check spelling mistakes.
Return corrected text and list the spelling corrections.
""",

                "Rewrite":
                    """
Rewrite the text naturally without changing the meaning.
""",

                "Professional":
                    """
Rewrite the text professionally and clearly.
""",

                "Friendly":
                    """
Rewrite the text in a friendly and polite tone.
""",

                "Simple English":
                    """
Rewrite using simple beginner-friendly English.
""",

                "Shorten":
                    """
Make the text shorter while keeping the important meaning.
""",

                "Expand":
                    """
Expand the text with clearer explanation and useful details.
""",

                "Summarize":
                    """
Summarize the text and give the important key points.
""",

                "Improve Writing":
                    """
Improve grammar, vocabulary, clarity and overall flow.
""",

                "Translate":
                    f"""
Translate the text into {target}.
Keep the original meaning.
"""
            }

            with st.spinner(
                "Working..."
            ):

                st.session_state.writing_result = (
                    ask_ai(
                        [],
                        text,
                        instruction=instruction_map[
                            action
                        ]
                    )
                )

    if st.session_state.writing_result:

        st.divider()

        st.markdown(
            "### ✨ Result"
        )

        st.markdown(
            st.session_state.writing_result
        )

        st.download_button(
            "⬇️ Download Result",
            data=st.session_state.writing_result,
            file_name="writing_result.txt",
            mime="text/plain",
            use_container_width=True
        )


# ============================================================
# ROADMAP
# ============================================================
with roadmap_tab:
    st.subheader("🗺️ Learning Roadmap")
    done=set(completed_topics(username,language)); topics_now=COURSES[language]["topics"]
    st.progress(len(done)/len(topics_now) if topics_now else 0)
    st.write(f"**{len(done)}/{len(topics_now)} topics completed**")
    for i,t in enumerate(topics_now,1): st.markdown(f"**{i}. {t}** — {('✅ Completed' if t in done else ('▶️ Current' if t==topic else '⬜ Not completed'))}")
    st.divider(); st.markdown("### 🤖 AI Study Plan")
    if st.button("Generate My 7-Day Plan",key="study_plan_btn"):
        st.markdown(generate_with_fallback(f"Create a practical 7-day beginner study plan for {language}. Completed: {list(done)}. Remaining: {topics_now}. Give concise daily tasks."))

# ============================================================
# CODE LAB
# ============================================================
with codelab_tab:
    st.subheader("🧑‍💻 AI Code Lab")
    mode=st.selectbox("Choose Tool",["Generate Code","Explain Code","Debug Code","Optimize Code","Convert Code","Generate Test Cases"],key="lab_mode")
    code=st.text_area("Paste code / describe requirement",height=250,key="lab_code")
    target=st.selectbox("Target Language",list(COURSES.keys()),index=list(COURSES.keys()).index(language) if language in COURSES else 0,key="lab_target")
    if st.button("🚀 Run Code Lab",key="run_code_lab",use_container_width=True):
        tasks={"Generate Code":"Generate clean beginner-friendly code and explain it step by step.","Explain Code":"Explain this code line by line and summarize the logic.","Debug Code":"Find bugs, explain them, provide corrected code and explain the fixes.","Optimize Code":"Improve readability and practical performance, then show the improved code.","Convert Code":f"Convert this code to {target} and explain major changes.","Generate Test Cases":"Create normal, edge and invalid test cases with expected outputs."}
        with st.spinner("AI is working..."): st.markdown(generate_with_fallback(f"You are a programming tutor. {tasks[mode]}\nLanguage: {target}\nInput:\n{code}"))

# ============================================================
# INTERVIEW
# ============================================================
with interview_tab:
    st.subheader("🎤 Interview Preparation")
    level=st.selectbox("Level",["Beginner","Intermediate","Advanced"],key="interview_level")
    countq=st.slider("Questions",3,10,5,key="interview_count")
    if "interview_questions" not in st.session_state: st.session_state.interview_questions=[]
    if st.button("Generate Interview",key="gen_interview",use_container_width=True):
        try: st.session_state.interview_questions=safe_json(generate_with_fallback(f"Create {countq} {level} {language} interview questions. Return JSON array with question, expected_points, sample_answer."))
        except Exception as e: st.error(f"Could not parse questions: {e}")
    if st.session_state.interview_questions:
        answers=[]
        for i,q in enumerate(st.session_state.interview_questions):
            qt=q.get("question",str(q)) if isinstance(q,dict) else str(q); st.markdown(f"**Q{i+1}. {qt}**"); answers.append(st.text_area("Your answer",key=f"int_ans_{i}",height=100))
        if st.button("🧠 Evaluate Answers",key="eval_interview",use_container_width=True):
            joined="\n\n".join(f"Q{i+1}: {a}" for i,a in enumerate(answers)); st.markdown(generate_with_fallback(f"Evaluate these {language} {level} interview answers. Give per-question feedback, missing concepts and an overall percentage.\n{joined}"))

# ============================================================
# NOTES
# ============================================================
with notes_tab:
    st.subheader("📝 My Notes")
    title=st.text_input("Note title",value=f"{language} - {topic}",key="note_title")
    content=st.text_area("Write your notes",height=220,key="note_content")
    if st.button("💾 Save Note",key="save_note",use_container_width=True):
        if content.strip(): save_note(username,title.strip() or "Untitled",content.strip(),language,topic); st.success("Note saved!")
        else: st.warning("Write something before saving.")
    for n in load_notes(username):
        with st.expander(f"📌 {n['title']} — {n['language']} / {n['topic']}"):
            st.write(n["content"]); st.download_button("⬇️ Download",n["content"],file_name=f"{n['title']}.txt",key=f"note_dl_{n['id']}")
            if st.button("🗑 Delete",key=f"note_del_{n['id']}"): delete_note(n["id"],username); st.rerun()

# ============================================================
# CERTIFICATE
# ============================================================
with certificate_tab:
    st.subheader("🏆 Course Certificate")
    cert_course=st.selectbox("Certificate Course",list(COURSES.keys()),key="cert_course")
    cert_topics=COURSES[cert_course]["topics"]; cert_done=completed_topics(username,cert_course); st.progress(len(cert_done)/len(cert_topics) if cert_topics else 0); st.write(f"**{len(cert_done)}/{len(cert_topics)} topics completed**")
    if len(cert_done)==len(cert_topics) and cert_topics:
        st.success("🎉 Course completion requirement met!")
        if st.button("📜 Generate Certificate PDF",key="cert_btn",use_container_width=True):
            try:
                from reportlab.lib.pagesizes import landscape,A4
                from reportlab.pdfgen import canvas
                from reportlab.lib.units import mm
                buf=io.BytesIO(); c=canvas.Canvas(buf,pagesize=landscape(A4)); w,h=landscape(A4)
                c.setFont("Helvetica-Bold",28); c.drawCentredString(w/2,h-55*mm,"INFINITY BOT"); c.setFont("Helvetica-Bold",22); c.drawCentredString(w/2,h-75*mm,"Certificate of Completion"); c.setFont("Helvetica",15); c.drawCentredString(w/2,h-95*mm,"This certifies that"); c.setFont("Helvetica-Bold",22); c.drawCentredString(w/2,h-112*mm,username); c.setFont("Helvetica",15); c.drawCentredString(w/2,h-130*mm,f"has completed the {cert_course} learning path."); c.setFont("Helvetica",11); c.drawCentredString(w/2,25*mm,f"Issued: {datetime.now().strftime('%d %B %Y')} • Infinity Bot"); c.save()
                st.download_button("⬇️ Download Certificate",buf.getvalue(),file_name=f"{username}_{cert_course}_certificate.pdf",mime="application/pdf")
            except ImportError: st.error("reportlab is not installed. Run: pip install reportlab")
    else: st.info("Complete all topics in this course to unlock the certificate.")


# ============================================================
# PROGRESS
# ============================================================

with progress_tab:

    st.subheader(
        "📊 My Progress"
    )

    completed = completed_topics(
        username,
        language
    )

    total = len(
        COURSES[language]["topics"]
    )

    count = len(
        completed
    )

    percentage = (
        count / total * 100
        if total
        else 0
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        st.metric(
            "Completed",
            count
        )

    with c2:

        st.metric(
            "Total Topics",
            total
        )

    with c3:

        st.metric(
            "Progress",
            f"{percentage:.0f}%"
        )

    st.progress(
        percentage / 100
    )

    st.divider()

    st.markdown(
        f"### 📚 {language}"
    )

    for item in COURSES[
        language
    ]["topics"]:

        if item in completed:

            st.success(
                f"✅ {item}"
            )

        else:

            st.write(
                f"⬜ {item}"
            )

    st.divider()

    # --------------------------------------------------------
    # QUIZ HISTORY
    # --------------------------------------------------------

    st.markdown(
        "### 🧪 Quiz History"
    )

    scores = db.execute(
        """
        SELECT topic,
               score_pct,
               correct_count,
               total_count,
               timestamp
        FROM lms_scores_v3
        WHERE username=?
        ORDER BY id DESC
        """,
        (
            username,
        )
    ).fetchall()

    if scores:

        df = pd.DataFrame(
            [
                dict(row)
                for row in scores
            ]
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        try:

            st.bar_chart(
                df.set_index(
                    "topic"
                )["score_pct"]
            )

        except Exception:

            pass

    else:

        st.info(
            "No quiz attempts yet."
        )

    st.divider()

    # --------------------------------------------------------
    # ACHIEVEMENTS
    # --------------------------------------------------------

    st.markdown(
        "### 🏆 Achievements"
    )

    if count >= 1:

        st.success(
            "🌱 First topic completed"
        )

    if count >= 5:

        st.success(
            "🔥 5 topics completed"
        )

    if count >= 10:

        st.success(
            "🚀 10 topics completed"
        )

    if count >= 16:

        st.success(
            "🏆 Python course completed"
        )

    if not completed:

        st.info(
            "Complete your first topic to unlock achievements."
        )