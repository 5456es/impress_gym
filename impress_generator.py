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
    TEXT_FORMATTING_TEXTBOX = "text_formatting_textbox"
    INSERT_TABLE = "insert_table"
    DELETE_TEXT_TEXTBOX = "delete_text_textbox"
    
    
    TEXT_FORMATTING_CONTENT_ = "text_formatting_content"    
    REPLACE_TEXT_ = "replace_text"
    INSERT_NOTE_ = "insert_note"
    FULLFILL_TABLE_ = "fullfill_table"
    


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
            "select_box": """
            
            You are a task generator for LibreOffice Impress automation. Generate a realistic textbox selection task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST include the specific and direct description of the textbox to select",
                "content": {
                    "text_in_textbox": "The text in the textbox that needs to be selected",
                    "environment_excluding_the_target_textbox": {
                        "other_textboxes": ["Text in other textboxes that are not the target"](at most 2),
                        "background_color": "Color of the slide background(please ignore randomly 50/50)",
                    },
                },
                "expected_result": {
                    "verification_type": "textbox_selection",
                    "text_in_textbox": "The text in the textbox that needs to be selected",
                },
                "metadata": {
                    "scnario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            #The instruction should describe the TYPE of textbox to select, also give the specific text in it. Examples:
                - "Select the textbox that contains the email address 123456@abc.com"
                - "Select the textbox that contains the project description which is 'This project aims to improve...'"
                - "Select the conclusion paragraph that says 'In conclusion..'"
                - "Select the title of the slide which says 'Project Overview'"
            
            #The environment setup should follow the scenario, e.g., if the task is to select a textbox with a specific text, the environment should include other textboxes that are not the target. And also not in conflict with the instruction. Example:
                - scnario: "select the textbox that contains the conclusion paragraph summarizing the project impact which is 'In conclusion...'"
                - "other_textboxes": [
                    "Introduction: This presentation outlines our green building initiative launched in Q1.",
                    "Objectives: Reduce energy consumption by 30percent over two years."
                ],
                - "text_in_textbox": "In conclusion, the project successfully reduced energy usage by 32%, exceeding our initial goals and demonstrating the effectiveness of our sustainability model."
            
            #The text in textbox don't be too long, just a few sentences is enough.
            
            """,
            
            "select_content" : """
            
            You are a task generator for LibreOffice Impress automation. Generate a realistic content selection task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify what text to select",
                "content": {
                    "target_text": "The specific text to select - appropriate length for the use case",
                    "full_text": "The full text in the textbox where the target text is located",
                },
                "expected_result": {
                    "verification_type": "text_selection",
                    "target_text": "The specific text to select - appropriate length for the use case"
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            The instruction should also describe the TYPE of content to select, not just giving the specific text. Examples:
                - "Select the email address 'contact@company.com' in the document"
                - "Highlight the entire paragraph about the project timeline"
                - "Select the phone number '(555) 123-4567' in the contact section"
                
            The full text should be realistic and not too long, just a few sentences is enough.
            """,
            
            
            "text_formatting_textbox": """
            You are a task generator for LibreOffice Impress automation. Generate a realistic text formatting task.
        
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify which textbox to apply format and how",
                "content": {
                    "textbox": "The description of the textbox to apply format - MUST include the text in it",
                    "formatting": "MUST be a JSON object with EXACTLY ONE property. Examples: {\"bold\": true} OR {\"font_size\": 16} OR {\"font\": \"Arial\"} OR {\"color\": \"0xFF0000\"} OR {\"strikethrough\": true} OR {\"alignment\": \"center\"} "
                },
                "expected_result": {
                    "verification_type": "has_formatting",
                    "expected_formatting": "same as formatting above"
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            The instruction should describe the TYPE of textbox to apply format, also give the specific text in it. Examples:
                - "Apply bold formatting to the textbox that contains the project description 'This project aims to improve...'"
                - "Change the font size to 16 for the textbox that contains the conclusion paragraph 'In conclusion...'"
                - "Set the font color to red in textbox that contains the title 'Project Overview'"
                - "Apply strikethrough to the textbox that contains the email address '
                - "Apply center alignment to the textbox that contains the phone number '(555) 123-4567'"
            
            Apply ONLY ONE formatting property per task. Examples:
            - "Make the title 'Quarterly Report - Q4 2024' bold" 
            → formatting: {"bold": true}
            - "Change the font size of the 'Executive Summary' to 16"
            → formatting: {"font_size": 16}
            - "Change the color of the font in the textbox 'Important Notice' to red"
            → formatting: {"color": "0xFF0000"}
            - "Strike through the text 'Important Notice'" 
            → formatting: {"strikethrough": true}
            - "Change the font of 'Executive Summary' to Arial" 
            → formatting: {"font": "Arial"}
            - "Center align the text in the textbox 'Contact Information'"
            → formatting: {"alignment": "center"}
            
            Consider aiming to apply the formatting to the entire textbox, not just a part of it. So when you describe the textbox, it should be clear that the formatting applies to the whole textbox. If 
            you don't mention the 'textbox' but the content, it should be the full content of the textbox.
            """,
            
            "insert_table": """
            
            You are a task generator for LibreOffice Impress automation. Generate a realistic blank table insertion task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify the rows and columns of the table to insert",
                "content": {
                    "table_structure": {
                        "rows": "Number of rows in the table to insert",
                        "columns": "Number of columns in the table to insert"
                    },                    
                },
                "expected_result": {
                    "verification_type": "table_insertion",
                    "table_structure": {
                        "rows": "Number of rows in the table to insert",
                        "columns": "Number of columns in the table to insert"
                    }
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            Only instruct the model to insert a blank table, not to fill it with any data. The instruction should specify the number of rows and columns in the table. Examples:
                - "Insert a table with 3 rows and 4 columns"
                - "Create a table with 5 rows and 2 columns for data entry"
                - "Add a table with 2 rows and 3 columns to the slide"
            """,
            
            "delete_text_textbox": """
            You are a task generator for LibreOffice Impress automation. Generate a realistic text deletion task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify the text to delete",
                "content": {
                    "text_to_delete": "The specific text to delete - shoule be full text in the textbox, appropriate length for the use case"
                },
                "expected_result": {
                    "verification_type": "text_deletion",
                    "deleted_text": "The specific text that was deleted - should be the same as text_to_delete"
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            The instruction should describe the TYPE of text to delete, also give the specific text in it. Examples:
                - "Delete the email address"
                - "Delete the project description 'This project aims to improve...'"
                - "Delete the conclusion paragraph 'In conclusion...'"
                - "Delete the title 'Project Overview'"
                - "Delete the phone number '(555) 123-4567'"
            
            The text to delete should be the full text in the textbox, appropriate length for the use case. It should not be too long. Also the full text to delete should be specified in the instruction, not just the specific text to delete.
            """
        }
