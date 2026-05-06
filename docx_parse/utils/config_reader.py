import json
import os

CONFIG_FILE_NAME = os.getenv("DOCX_PARSE_CONFIG_JSON", "docx_parse.json")


def read_config():
    if os.path.isabs(CONFIG_FILE_NAME):
        config_file = CONFIG_FILE_NAME
    else:
        config_file = os.path.join(os.path.expanduser("~"), CONFIG_FILE_NAME)
    if not os.path.exists(config_file):
        return None
    with open(config_file, "r", encoding="utf-8") as f:
        return json.load(f)


def get_latex_delimiter_config():
    config = read_config()
    if not config:
        return None
    return config.get("latex-delimiter-config")
