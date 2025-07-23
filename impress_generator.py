import json
import random
import requests
import shlex
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

class TaskType(Enum):
    # 1. 选框/Select_Box
    # 2. 选文字/Select_Content
    # 3. 修改文字格式/Text_Formatting
    # 1. 大小
    # 2. 加粗
    # 3. 删除线
    # 4. 对齐
    # 5. 颜色
    # 4. 插入表格/Insert_Table
    # 5. 删除文字/Delete_Text
    # 6. 更换文字/Replace_Text
    SELECT_BOX = "select_box"
    SELECT_CONTENT = "select_content"
    TEXT_FORMATTING = "text_formatting"
    INSERT_TABLE = "insert_table"
    DELETE_TEXT = "delete_text"
    REPLACE_TEXT = "replace_text"
    


@dataclass
class TaskData:
    """LLM生成的任务数据结构"""
    instruction: str
    content: Dict[str, Any]
    expected_result: Dict[str, Any]
    metadata: Dict[str, Any]
    
class FullLLMTaskGenerator:
    """完全由LLM驱动的任务生成器"""
    
    def __init__(self, api_key: str, model: str = "gpt-4o", base_url: str = "https://api.openai.com/v1/chat/completions"):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        
        # 直接提示版本：指令中包含具体内容
        self.direct_prompts = {
            "select_box":"""
            
            You are a task generator for LibreOffice Impress automation. Generate a realistic textbox selection task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST include the specific description of the textbox to select",
                "content": {
                    "text_in_textbox": "The text in the textbox that needs to be selected",
                    "environment_excluding_the_target_textbox": {
                        "other_textboxes": ["Text in other textboxes that are not the target"](at most 2),
                        "background_color": "Color of the slide background(please ignore randomly 50/50)",
                    },
                },
                "expected_result": {
                    "verification_type": "contains_text",
                    "text_in_textbox": "The text in the textbox that needs to be selected",
                },
                "metadata": {
                    "scnario": "brief description of use case, e.g., 'select the textbox that contasins the summary of another textbox'",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            The instruction should describe the TYPE of textbox to select, not just giving the specific text in it. Examples:
                - "Select the email address"
                - "Select the textbox that contains the project description"
                - "Select the phone number in this slide"
                - "Select the conclusion paragraph"
            
            The environment setup should follow the scenario, e.g., if the task is to select a textbox with a specific text, the environment should include other textboxes that are not the target. And also not in conflict with the instruction. Example:
                - scnario: "select the textbox that contains the conclusion paragraph summarizing the project impact"
                - "other_textboxes": [
                    "Introduction: This presentation outlines our green building initiative launched in Q1.",
                    "Objectives: Reduce energy consumption by 30percent over two years."
                ],
                - "text_in_textbox": "In conclusion, the project successfully reduced energy usage by 32%, exceeding our initial goals and demonstrating the effectiveness of our sustainability model."

            """,
        }
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            """
        