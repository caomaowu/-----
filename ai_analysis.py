import os
import base64
import json
import time
from openai import OpenAI
from utils import log

class AIAnalyzer:
    def __init__(self, config):
        self.config = config.get("ai", {})
        self.enabled = self.config.get("enabled", False)
        self.base_url = self.config.get("base_url", "https://api.openai.com/v1")
        self.api_key_env = self.config.get("api_key_env", "OPENAI_API_KEY")
        self.active_model = self.config.get("active_model", "gpt-4o-mini")
        self.client = None
        
        if self.enabled:
            # Priority 1: Try to get from environment variable
            api_key = os.environ.get(self.api_key_env)
            
            # Priority 2: If not found in env, check if the config value itself looks like a key
            # (User might have pasted the key directly into the config file instead of the env var name)
            if not api_key:
                if self.api_key_env.startswith("sk-") or len(self.api_key_env) > 20:
                    api_key = self.api_key_env
                    log("Note: Using API key directly from config file.")
            
            if not api_key:
                log(f"Warning: {self.api_key_env} environment variable not set. AI analysis might fail.")
            
            try:
                self.client = OpenAI(
                    base_url=self.base_url,
                    api_key=api_key
                )
                log(f"AI Analyzer initialized with model {self.active_model}")
            except Exception as e:
                log(f"Error initializing OpenAI client: {e}")
                self.enabled = False

    def encode_image(self, image_path):
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode('utf-8')
        except Exception as e:
            log(f"Error encoding image {image_path}: {e}")
            return None

    def analyze_image(self, image_path, key):
        if not self.client:
            return None

        base64_image = self.encode_image(image_path)
        if not base64_image:
            return None

        prompt = f"""
        你是一个专业的工程报告助手。请分析这张图片（Key: {key}）。
        请输出严格的 JSON 格式，包含以下字段：
        - summary: 一句话结论（40字以内）
        - bullets: 关键点列表（2-4条）
        
        不要输出 JSON 以外的任何内容。
        """

        try:
            response = self.client.chat.completions.create(
                model=self.active_model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{base64_image}"
                                },
                            },
                        ],
                    }
                ],
                max_tokens=500,
                temperature=0.2,
            )
            
            content = response.choices[0].message.content
            # Try to parse JSON
            content_clean = content.strip()
            if content_clean.startswith("```"):
                import re
                content_clean = re.sub(r"^```json\s*", "", content_clean)
                content_clean = re.sub(r"^```\s*", "", content_clean)
                content_clean = re.sub(r"\s*```$", "", content_clean)
            
            return json.loads(content_clean)
            
        except Exception as e:
            log(f"Error calling AI API for {key}: {e}")
            return None

    def format_result(self, result):
        if not result:
            return "AI分析失败。"
        
        try:
            summary = result.get("summary", "")
            bullets = result.get("bullets", [])
            
            text = f"{summary}\n"
            for b in bullets:
                text += f"• {b}\n"
            return text.strip()
        except:
            return str(result)

    def process_presentation(self, ppt, temp_dir, local_images_map=None):
        if not self.enabled:
            return

        log("Starting AI Analysis phase...")
        triggers = ppt.scan_ai_triggers()
        log(f"Found {len(triggers)} AI triggers.")
        
        if not triggers:
            return

        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
            
        local_images_map = local_images_map or {}
        
        for trigger in triggers:
            key = trigger["key"]
            shape = trigger["shape"]
            full_match = trigger["full_match_str"]
            
            log(f"Processing AI trigger for key: {key}")
            
            # Determine image source: Local File (Priority) > Export from PPT
            target_image_path = None
            
            # 1. Try local map
            if key in local_images_map and os.path.exists(local_images_map[key]):
                target_image_path = local_images_map[key]
                log(f"Using local image source: {target_image_path}")
            
            # 2. Try export from PPT if not found locally
            if not target_image_path:
                export_path = os.path.join(temp_dir, f"ai_temp_{key}.png")
                if ppt.export_smarttag_image(key, export_path):
                    target_image_path = export_path
                    log(f"Exported image from PPT: {target_image_path}")
            
            if target_image_path:
                # Analyze
                result = self.analyze_image(target_image_path, key)
                
                # Format
                if result:
                    new_text = self.format_result(result)
                else:
                    new_text = "AI分析失败：无法获取结果"
                
                # Replace
                ppt.replace_text_content(shape, full_match, new_text)
                log(f"Updated text for {key}")
                
            else:
                log(f"Target image for {key} not found (Locally or in PPT).")
                ppt.replace_text_content(shape, full_match, "未找到目标图片。")
