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
    INSERT_RESIZE_IMAGE = "insert_resize_image"

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

    def __init__(
        self,
        api_key: str,
        model: str = "gpt-4o",
        base_url: str = "https://api.openai.com/v1/chat/completions",
    ):
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
            "select_content": """
            
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
                - "Highlight the entire paragraph about the project timeline by selecting the text"
                - "Select the phone number '(555) 123-4567' in the contact section"
                
            The full text should be realistic and not too long, just a few sentences is enough.
            """,
            "text_formatting_textbox": """
            You are a task generator for LibreOffice Impress automation. Generate a realistic text formatting task.
        
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify which textbox to apply format and how",
                "content": {
                    "text_in_target_textbox": "The exact full text in the textbox where the formatting should be applied - appropriate length for the use case",
                    "formatting": "MUST be a JSON object with EXACTLY ONE property. Examples: {\"bold\": true} OR {\"font_size\": 16} OR {\"font\": \"Arial\"} OR {\"color\": \"0xFF0000\"} OR {\"strikethrough\": true} OR {\"alignment\": \"center\"} "
                },
                "expected_result": {
                    "verification_type": "has_formatting",
                    "text_in_target_textbox": "The exact full text in the textbox after formatting - should be the same as text_in_target_textbox",
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
            you either mention the 'textbox' or point out the full content of the textbox.
            """,
            "insert_table": """
            
            You are a task generator for LibreOffice Impress automation. Generate a realistic blank table insertion task.
            
            Return ONLY a valid JSON object with this exact structure:
            {
                "instruction": "Natural language instruction for the user - MUST specify the rows and columns of the table to insert",
                "content": {
                    "table_structure": {
                        "rows": "Number of rows in the table to insert(range 5 to 15, diverse)",
                        "columns": "Number of columns in the table to insert(range 5 to 15)"
                    },                    
                },
                "expected_result": {
                    "verification_type": "table_insertion",
                    "table_structure": {
                        "rows": "Number of rows in the table to insert(range 5 to 15)",
                        "columns": "Number of columns in the table to insert(range 5 to 15)"
                    }
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            Only instruct the model to insert a blank table, not to fill it with any data. The instruction should specify the number of rows and columns in the table. Examples:
                - "Insert a table with 3 rows and 6 columns"
                - "Create a table with 7 rows and 5 columns for data entry"
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
            """,
            "insert_resize_image": """
            You are a task generator for LibreOffice Impress automation. Generate a realistic image insertion and resizing task.
            
            Return ONLY a valid JSON object with this exact structure:
            
            {
                "instruction": "Natural language instruction for the user - MUST specify the image to insert and how to resize it",
                "content": {
                    "image_path": "/home/user/Desktop/image_to_insert.jpg",  # Fixed path where the image is stored
                    "resize_dimensions": {
                        "width": "Width to resize the image to (in integer cms)",
                        "height": "Height to resize the image to (in integer cms)"
                    }
                },
                "expected_result": {
                    "verification_type": "image_insertion_and_resizing",
                    "image_path": "/home/user/Desktop/image_to_insert.jpg",  # Fixed path where the image is stored
                    "resize_dimensions": {
                        "width": "Width of the resized image (in interger cms)",
                        "height": "Height of the resized image (in interger cms)"
                    }
                },
                "metadata": {
                    "scenario": "brief description of use case",
                    "difficulty": "easy|medium|hard"
                }
            }
            
            IMPORTANT:
            Due to the limitations of the current API, you can only specify the image path and the dimensions to resize it to. 
            Also the image path is fixed which is /home/user/Desktop/image_to_insert.jpg.
            
            Since we don't actually know the content of the image, so don't include any specific content in the instruction. Just focus on the insertion and resizing of the image for some general use cases.
            """,
        }

        # 场景类别，用于生成不同背景的任务
        self.scenario_categories = [
            "business_presentation",  # 商务演讲、销售方案、季度总结
            "educational_lecture",  # 教学课件、PPT课程内容
            "academic_defense",  # 毕业答辩、研究成果展示
            "marketing_pitch",  # 产品发布、创业路演
            "project_overview",  # 项目计划、阶段汇报
            "team_meeting",  # 内部例会、进度通报
            "training_material",  # 培训课程、指导文档
            "conference_talk",  # 行业大会、专业分享
            "technical_demo",  # 技术原理、架构展示
            "status_update",  # 工作进展、日报周报
            "portfolio_showcase",  # 设计作品集、个人展示
            "event_invitation",  # 活动通知、会议邀请
            "company_profile",  # 企业介绍、文化展示
            "financial_summary",  # 财务报表、营收结构
            "product_showcase",  # 产品特点、功能亮点
            "infographic_summary",  # 可视化信息图、数据导图
            "public_speaking",  # 演讲比赛、讲稿支持
            "interactive_quiz",  # 教学测验、答题交互
            "client_proposal",  # 客户方案、商业计划书
            "internal_onboarding",  # 新员工培训、内部手册
        ]

    def call_llm(
        self, system_prompt: str, user_prompt: str, max_retries: int = 3
    ) -> Dict[str, Any]:
        """调用LLM API"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "max_tokens": 1500,  # 增加token限制以支持更长文本
            "temperature": 0.8,
        }

        for attempt in range(max_retries):
            try:
                response = requests.post(
                    self.base_url, headers=headers, json=data, timeout=60
                )  # 增加超时时间

                if response.status_code == 200:
                    result = response.json()
                    content = result["choices"][0]["message"]["content"].strip()

                    # 尝试解析JSON
                    try:
                        # 清理可能的markdown标记
                        if content.startswith("```json"):
                            content = content.split("```json")[1].split("```")[0]
                        elif content.startswith("```"):
                            content = content.split("```")[1].split("```")[0]

                        return json.loads(content)
                    except json.JSONDecodeError as e:
                        print(f"JSON parsing error (attempt {attempt + 1}): {e}")
                        print(f"Raw content: {content}")
                        if attempt == max_retries - 1:
                            raise Exception(
                                f"Failed to parse JSON after {max_retries} attempts"
                            )
                        continue
                else:
                    print(f"API error (attempt {attempt + 1}): {response.status_code}")
                    if attempt == max_retries - 1:
                        raise Exception(
                            f"API call failed with status {response.status_code}"
                        )

            except Exception as e:
                print(f"Request error (attempt {attempt + 1}): {e}")
                if attempt == max_retries - 1:
                    raise

    def generate_task_data(
        self,
        task_type: TaskType,
        scenario_category: str = None,
        direct_instruction_ratio: float = 1,
    ) -> TaskData:
        """生成完整的任务数据

        Args:
            task_type: 任务类型
            scenario_category: 场景类别
            direct_instruction_ratio: 直接指令比例 (0.0-1.0)，0.5表示50%直接指令，50%结构指令
        """
        if scenario_category is None:
            scenario_category = random.choice(self.scenario_categories)

        # 根据比例随机选择使用直接提示还是结构提示
        use_direct_prompt = random.random() < direct_instruction_ratio

        if use_direct_prompt:
            system_prompt = self.direct_prompts[task_type.value]
            instruction_type = "direct"
        else:
            system_prompt = self.structural_prompts[task_type.value]
            instruction_type = "structural"

        user_prompt = f"""Generate a {task_type.value} task for a {scenario_category} scenario. 
        Make it realistic, practical, and varied.
        
        IMPORTANT: Choose appropriate content length based on the realistic use case, especially not too long

        Focus on creating tasks that someone would actually need to do when working with real impress file."""

        try:
            llm_response = self.call_llm(system_prompt, user_prompt)

            # 将document_content添加到content字段中
            content = llm_response["content"].copy()

            return TaskData(
                instruction=llm_response["instruction"],
                content=content,
                expected_result=llm_response["expected_result"],
                metadata={
                    **llm_response.get("metadata", {}),
                    "scenario_category": scenario_category,
                    "generated_by_llm": True,
                    "instruction_type": instruction_type,
                },
            )
        except Exception as e:
            print(f"LLM generation failed for {task_type.value}: {e}")


class LibreOfficeImpressTaskGenerator:
    def __init__(
        self,
        llm_api_key: str,
        model: str = "gpt-4o",
        direct_instruction_ratio: float = 1,
    ):
        self.llm_generator = FullLLMTaskGenerator(llm_api_key, model)
        self.direct_instruction_ratio = direct_instruction_ratio

        self.base_config = [
            {
                "type": "launch",
                "parameters": {
                    "command": [
                        "libreoffice --impress --accept='socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager'"
                    ],
                    "shell": True,
                },
            },
            {"type": "sleep", "parameters": {"seconds": 5}},
            {
                "type": "execute",
                "parameters": {
                    "command": ["curl -X POST localhost:5011/api/connect"],
                    "shell": True,
                },
            },
        ]

    def generate_single_task(
        self,
        task_type: TaskType = None,
        scenario_category: str = None,
        direct_instruction_ratio: float = None,
    ) -> Dict[str, Any]:
        """生成单个任务"""
        if task_type is None:
            task_type = random.choice(list(TaskType))

        if direct_instruction_ratio is None:
            direct_instruction_ratio = self.direct_instruction_ratio

        task_data = self.llm_generator.generate_task_data(
            task_type, scenario_category, direct_instruction_ratio
        )
        task_id = f"{task_type.value}_{random.randint(1000, 9999)}"

        return self.create_task_from_llm_data(task_id, task_type, task_data)

    def create_task_from_llm_data(
        self, task_id: str, task_type: TaskType, task_data: TaskData
    ) -> Dict[str, Any]:
        """根据LLM生成的数据创建完整任务"""

        base_task = {
            "id": task_id,
            "snapshot": "libreoffice_impress",
            "instruction": task_data.instruction,
            "source": "",
            "task_type": "scaling",
            "config": self.base_config.copy(),
            "trajectory": "trajectories/",
            "related_apps": [],
            "metadata": task_data.metadata,
        }

        if task_type == TaskType.SELECT_BOX:
            return self._create_select_box_task(base_task, task_data)
        elif task_type == TaskType.SELECT_CONTENT:
            return self._create_select_content_task(base_task, task_data)
        elif task_type == TaskType.TEXT_FORMATTING_TEXTBOX:
            return self._create_text_formatting_task(base_task, task_data)
        elif task_type == TaskType.INSERT_TABLE:
            return self._create_insert_table_task(base_task, task_data)
        elif task_type == TaskType.DELETE_TEXT_TEXTBOX:
            return self._create_delete_text_task(base_task, task_data)
        elif task_type == TaskType.INSERT_RESIZE_IMAGE:
            return self._create_insert_resize_image_task(base_task, task_data)

    def _create_select_box_task(
        self, base_task: Dict[str, Any], task_data: TaskData
    ) -> Dict[str, Any]:
        """创建选框任务"""

        expected = task_data.expected_result

        # {
        #     "instruction": "Select the textbox that contains the slide's key takeaway message which is 'The merger will increase our market share by 15%.'",
        #     "content": {
        #         "text_in_textbox": "The merger will increase our market share by 15%.",
        #         "environment_excluding_the_target_textbox": {
        #             "other_textboxes": [
        #                 "Introduction: We are excited to discuss the upcoming merger and its benefits.",
        #                 "Financial Overview: The merger is projected to bring a 10% increase in revenue."
        #             ],
        #             "background_color": "blue"
        #         }
        #     },
        #     "expected_result": {
        #         "verification_type": "textbox_selection",
        #         "text_in_textbox": "The merger will increase our market share by 15%."
        #     },
        #     "metadata": {
        #         "scnario": "Select the key takeaway message in a business presentation about a merger.",
        #         "difficulty": "medium",
        #         "scenario_category": "business_presentation",
        #         "generated_by_llm": true,
        #         "instruction_type": "direct"
        #     }
        # }

        ### 1. 先添加页面，再删除到只剩一张
        add_cmd = (
            "curl -X POST http://localhost:5011/api/slide/new "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        # 构造 DELETE 命令
        delete_cmd = (
            "curl -X DELETE http://localhost:5011/api/slide/0 "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        setup_impress_command = [
            {"type": "execute", "parameters": {"command": [add_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
            {"type": "execute", "parameters": {"command": [delete_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
        ]
        base_task["config"].extend(setup_impress_command)

        ### 2. set up 框
        target_textbox = task_data.content["text_in_textbox"]
        env_textboxes = task_data.content[
            "environment_excluding_the_target_textbox"
        ].get("other_textboxes", [])

        add_textbox_config = []
        add_textbox_config.append(
            {
                "text": target_textbox,
                "x": random.randint(1000, 18000),
                "y": random.randint(1000, 14000),
                "formatting": {
                    "bold": random.choice([True, False]),
                    "italic": random.choice([True, False]),
                    "font_size": random.randint(10, 50),
                    "alignment": random.choice(["left", "right", "center"]),
                },
            }
        )
        for text in env_textboxes:
            add_textbox_config.append(
                {
                    "text": text,
                    "x": random.randint(1000, 18000),
                    "y": random.randint(1000, 14000),
                    "formatting": {
                        "bold": random.choice([True, False]),
                        "italic": random.choice([True, False]),
                        "font_size": random.randint(10, 50),
                        "alignment": random.choice(["left", "right", "center"]),
                    },
                    "width": random.randint(8000, 12000),
                    "height": random.randint(1500, 4000),
                }
            )
        for textbox_config in add_textbox_config:
            add_text_cmd = (
                "curl -X POST http://localhost:5011/api/slide/add-text "
                "-H 'Content-Type: application/json' "
                f"-d {shlex.quote(json.dumps(textbox_config))}"
            )
            base_task["config"].append(
                {
                    "type": "execute",
                    "parameters": {"command": [add_text_cmd], "shell": True},
                }
            )
            base_task["config"].append(
                {"type": "sleep", "parameters": {"seconds": 1}},
            )

        # 3. 设置验证
        base_task["evaluator"] = {
            "postconfig": [
                {
                    "type": "execute",
                    "parameters": {
                        "command": [
                            "python",
                            "-c",
                            "import pyautogui; import time; pyautogui.press('delete'); time.sleep(0.5);",
                        ]
                    },
                }
            ],
            "func": "textbox_selection_verification",
            "result": {
                "type": "current_content",
                "verification": expected["verification_type"],
            },
            "expected": {
                "type": "rule",
                "rules": {
                    "other_textboxes": task_data.content[
                        "environment_excluding_the_target_textbox"
                    ]["other_textboxes"],
                },
            },
        }
        return base_task

    def _create_select_content_task(
        self, base_task: Dict[str, Any], task_data: TaskData
    ) -> Dict[str, Any]:
        """创建选内容任务"""
        expected = task_data.expected_result

        # {
        #     "instruction": "Natural language instruction for the user - MUST specify what text to select",
        #     "content": {
        #         "target_text": "The specific text to select - appropriate length for the use case",
        #         "full_text": "The full text in the textbox where the target text is located",
        #     },
        #     "expected_result": {
        #         "verification_type": "text_selection",
        #         "target_text": "The specific text to select - appropriate length for the use case"
        #     },
        #     "metadata": {
        #         "scenario": "brief description of use case",
        #         "difficulty": "easy|medium|hard"
        #     }
        # }

        add_cmd = (
            "curl -X POST http://localhost:5011/api/slide/new "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        # 构造 DELETE 命令
        delete_cmd = (
            "curl -X DELETE http://localhost:5011/api/slide/0 "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        setup_impress_command = [
            {"type": "execute", "parameters": {"command": [add_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
            {"type": "execute", "parameters": {"command": [delete_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
        ]
        base_task["config"].extend(setup_impress_command)

        ### 2. set up 框
        full_text = task_data.content["full_text"]
        target_text = task_data.content["target_text"]

        add_textbox_config = []
        add_textbox_config.append(
            {
                "text": full_text,
                "x": random.randint(1000, 16000),
                "y": random.randint(1000, 10000),
                "formatting": {
                    "bold": random.choice([True, False]),
                    "italic": random.choice([True, False]),
                    "font_size": random.randint(10, 50),
                    "alignment": random.choice(["left", "right", "center"]),
                    "width": random.randint(8000, 12000),
                    "height": random.randint(1500, 4000),
                },
            }
        )

        for textbox_config in add_textbox_config:
            add_text_cmd = (
                "curl -X POST http://localhost:5011/api/slide/add-text "
                "-H 'Content-Type: application/json' "
                f"-d {shlex.quote(json.dumps(textbox_config))}"
            )
            base_task["config"].append(
                {
                    "type": "execute",
                    "parameters": {"command": [add_text_cmd], "shell": True},
                }
            )
            base_task["config"].append(
                {"type": "sleep", "parameters": {"seconds": 1}},
            )

        # 3. 设置验证
        base_task["evaluator"] = {
            "func": "content_selection_verification",
            "result": {
                "type": "selected_content",
                "verification": expected["verification_type"],
            },
            "expected": {
                "type": "rule",
                "rules": {
                    "target_text": task_data.content["target_text"],
                },
            },
        }
        return base_task

    def _create_text_formatting_task(
        self, base_task: Dict[str, Any], task_data: TaskData
    ) -> Dict[str, Any]:
        """创建文本格式化任务"""
        expected = task_data.expected_result

        # {
        #     "instruction": "Natural language instruction for the user - MUST specify which textbox to apply format and how",
        #     "content": {
        #         "text_in_target_textbox": "The exact full text in the textbox where the formatting should be applied - appropriate length for the use case",
        #         "formatting": "MUST be a JSON object with EXACTLY ONE property. Examples: {\"bold\": true} OR {\"font_size\": 16} OR {\"font\": \"Arial\"} OR {\"color\": \"0xFF0000\"} OR {\"strikethrough\": true} OR {\"alignment\": \"center\"} "
        #     },
        #     "expected_result": {
        #         "verification_type": "has_formatting",
        #         "text_in_target_textbox": "The exact full text in the textbox after formatting - should be the same as text_in_target_textbox",
        #         "expected_formatting": "same as formatting above"
        #     },
        #     "metadata": {
        #         "scenario": "brief description of use case",
        #         "difficulty": "easy|medium|hard"
        #     }
        # }

        # 1. 先添加页面，再删除到只剩一张

        add_cmd = (
            "curl -X POST http://localhost:5011/api/slide/new "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        # 构造 DELETE 命令
        delete_cmd = (
            "curl -X DELETE http://localhost:5011/api/slide/0 "
            "-H 'Content-Type: application/json' "
            f"-d {shlex.quote(json.dumps({}))}"
        )

        setup_impress_command = [
            {"type": "execute", "parameters": {"command": [add_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
            {"type": "execute", "parameters": {"command": [delete_cmd], "shell": True}},
            {"type": "sleep", "parameters": {"seconds": 5}},
        ]
        base_task["config"].extend(setup_impress_command)

        ### 2. set up 框
        target_textbox = task_data.content["text_in_target_textbox"]
        add_textbox_config = [
            {
                "text": target_textbox,
                "x": random.randint(1000, 16000),
                "y": random.randint(1000, 10000),
                "width": random.randint(8000, 12000),
                "height": random.randint(1500, 4000),
            }
        ]

        for textbox_config in add_textbox_config:
            add_text_cmd = (
                "curl -X POST http://localhost:5011/api/slide/add-text "
                "-H 'Content-Type: application/json' "
                f"-d {shlex.quote(json.dumps(textbox_config))}"
            )
            base_task["config"].append(
                {
                    "type": "execute",
                    "parameters": {"command": [add_text_cmd], "shell": True},
                }
            )
            base_task["config"].append(
                {"type": "sleep", "parameters": {"seconds": 1}},
            )

        ### 3. 设置验证
        base_task["evaluator"] = {
            "func": "text_formatting_verification",
            "result": {
                "type": "current_content",
                "verification": expected["verification_type"],
            },
            "expected": {
                "type": "rule",
                "rules": {
                    "text_in_target_textbox": task_data.content[
                        "text_in_target_textbox"
                    ],
                    "expected_formatting": task_data.content["formatting"],
                },
            },
        }
        return base_task

    def _create_insert_table_task(
        self, base_task: Dict[str, Any], task_data: TaskData
    ) -> Dict[str, Any]:
        """创建插入表格任务"""
        expected = task_data.expected_result

        # {
        #     "instruction": "Natural language instruction for the user - MUST specify the rows and columns of the table to insert",
        #     "content": {
        #         "table_structure": {
        #             "rows": "Number of rows in the table to insert(range 5 to 15, diverse)",
        #             "columns": "Number of columns in the table to insert(range 5 to 15)"
        #         },
        #     },
        #     "expected_result": {
        #         "verification_type": "table_insertion",
        #         "table_structure": {
        #             "rows": "Number of rows in the table to insert(range 5 to 15)",
        #             "columns": "Number of columns in the table to insert(range 5 to 15)"
        #         }
        #     },
        #     "metadata": {
        #         "scenario": "brief description of use case",
        #         "difficulty": "easy|medium|hard"
        #     }
        # }

        base_task["evaluator"] = {
            "func": "table_insertion_verification",
            "result": {
                "type": "current_content",
                "verification": expected["verification_type"],
            },
            "expected": {
                "type": "rule",
                "rules": {
                    "table_structure": task_data.content["table_structure"],
                },
            },
        }
        return base_task

    def _create_insert_resize_image_task(
        self, base_task: Dict[str, Any], task_data: TaskData
    ) -> Dict[str, Any]:
        """创建插入和调整大小的图片任务"""
        expected = task_data.expected_result

        # {
        #     "instruction": "Natural language instruction for the user - MUST specify the image to insert and how to resize it",
        #     "content": {
        #         "image_path": "/home/user/Desktop",  # Fixed path where the image is stored
        #         "resize_dimensions": {
        #             "width": "Width to resize the image to (in cms)",
        #             "height": "Height to resize the image to (in cms)"
        #         }
        #     },
        #     "expected_result": {
        #         "verification_type": "image_insertion_and_resizing",
        #         "image_path": "/home/user/Desktop",  # Fixed path where the image is stored
        #         "resize_dimensions": {
        #             "width": "Width of the resized image (in cms)",
        #             "height": "Height of the resized image (in cms)"
        #         }
        #     },
        #     "metadata": {
        #         "scenario": "brief description of use case",
        #         "difficulty": "easy|medium|hard"
        #     }
        # }
        #     "config": [
        #     {
        #         "type": "download",
        #         "parameters": {
        #             "files": [
        #                 {
        #                     "url": "https://agent-files.deva.msh.team/osworld/benchmark_files/libreoffice_impress/0f84bef9-9790-432e-92b7-eece357603fb_multimedia_classroom_podium-2020.pptx",
        #                     "path": "/home/user/Desktop/multimedia_classroom_podium-2020.pptx"
        #                 }
        #             ]
        #         }
        #     },
        #     {
        #         "type": "open",
        #         "parameters": {
        #             "path": "/home/user/Desktop/multimedia_classroom_podium-2020.pptx"
        #         }
        #     }
        # ],
        # 1. 准备把图片先上传到指定路径

        image_file = f"https://agent-files.deva.msh.team/osworld/scaling_files/libreoffice_impress_gym_images/impress_gym_images/{random.randint(1, 100)}.jpg"

        upload_image_cmd = {
            "type": "download",
            "parameters": {
                "files": [
                    {
                        "url": image_file,
                        "path": "/home/user/Desktop/image_to_insert.jpg",
                    }
                ]
            },
        }
        base_task["config"].append(upload_image_cmd)

        # 2. 准备验证
        base_task["evaluator"] = {
            "func": "image_insertion_and_resizing_verification",
            "result": {
                "type": "current_content",
                "verification": expected["verification_type"],
            },
            "expected": {
                "type": "rule",
                "rules": {
                    "image_path": task_data.content["image_path"],
                    "resize_dimensions": task_data.content["resize_dimensions"],
                },
            },
        }

        return base_task


if __name__ == "__main__":
    # 示例用法
    with open("api_key.txt", "r") as f:
        api_key = f.read().strip()
    generator = LibreOfficeImpressTaskGenerator(
        llm_api_key=api_key,
        model="gpt-4o",
        direct_instruction_ratio=1,
    )
    task = generator.generate_single_task(
        task_type=TaskType.INSERT_TABLE,
        scenario_category="business_presentation",
    )
    print(json.dumps(task, indent=2, ensure_ascii=False))
    with open("test_tasks/insert_table.json", "w", encoding="utf-8") as f:
        json.dump(task, f, indent=2, ensure_ascii=False)
    # 生成的任务将包含完整的配置和指令
