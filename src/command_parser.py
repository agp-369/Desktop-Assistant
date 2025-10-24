from fuzzywuzzy import process

# Define the commands and their variations
COMMANDS = {
    "open_app": {
        "keywords": ["open", "launch", "start"],
        "min_ratio": 80
    },
    "close_app": {
        "keywords": ["close", "exit", "terminate", "quit"],
        "min_ratio": 80
    },
    "open_file": {
        "keywords": ["open file", "show file"],
        "min_ratio": 85
    },
    "play_youtube": {
        "keywords": ["play on youtube", "find on youtube", "search on youtube"],
        "min_ratio": 85
    },
    "type_text": {
        "keywords": ["type", "write", "enter"],
        "min_ratio": 80
    },
    "exit": {
        "keywords": ["exit", "goodbye", "quit", "shutdown"],
        "min_ratio": 90
    }
}

def parse_command(text):
    """
    Parses a command from the user's speech.
    Returns the command and the extracted arguments.
    """
    if not text:
        return None, None

    # Find the best matching command
    best_match_command = None
    best_match_ratio = 0
    best_match_keyword = None

    for command, data in COMMANDS.items():
        result = process.extractOne(text, data["keywords"])
        if result and result[1] > best_match_ratio and result[1] >= data["min_ratio"]:
            best_match_ratio = result[1]
            best_match_command = command
            best_match_keyword = result[0]

    if not best_match_command:
        return None, None

    # Extract the arguments from the text
    args_text = text.replace(best_match_keyword, "").strip()

    return best_match_command, args_text
