import deepl
import openai
from deep_translator import GoogleTranslator
import json
import os
from dotenv import load_dotenv

load_dotenv()

class TranslationService:
    def __init__(self, service_type="DeepL", api_key=None):
        self.service_type = service_type
        self.api_key = api_key
        self.glossary = self._load_glossary()
        self.cache = {} # In-memory cache for the current session

    def _load_glossary(self):
        try:
            with open("glossary.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def translate(self, text):
        if not text or not text.strip():
            return text
        
        # Check cache
        if text in self.cache:
            return self.cache[text]
        
        # Simple check if text is already Korean (contains Hangul)
        if any('\uac00' <= char <= '\ud7a3' for char in text):
            self.cache[text] = text
            return text
        
        result = text
        if self.service_type == "DeepL":
            result = self._translate_deepl(text)
        elif self.service_type == "OpenAI":
            result = self._translate_openai(text)
        elif self.service_type == "Free (Google)":
            result = self._translate_free(text)
        elif self.service_type == "Smart (OpenAI -> Free)":
            # Primary: OpenAI
            result = self._translate_openai(text)
            # If OpenAI fails (returns original text or None/Error), try Free
            if result is None or result == text:
                result = self._translate_free(text)
        
        # FINAL GUARD: Ensure we NEVER return None
        if result is None:
            result = text
            
        self.cache[text] = result
        return result

    def _translate_free(self, text):
        import time
        max_retries = 3
        for i in range(max_retries):
            try:
                # GoogleTranslator from deep-translator often works without a key
                result = GoogleTranslator(source='en', target='ko').translate(text)
                if result:
                    return result
                print(f"Free Translator Attempt {i+1} returned empty result.")
            except Exception as e:
                print(f"Free Translator Attempt {i+1} failed: {e}")
            
            if i < max_retries - 1:
                time.sleep(1)
        return text

    def _translate_deepl(self, text):
        if not self.api_key:
            return text # Just return original if no key
        try:
            translator = deepl.Translator(self.api_key)
            result = translator.translate_text(text, target_lang="KO")
            return result.text
        except Exception as e:
            print(f"DeepL Error: {e}")
            return text # Fallback to original text on error

    def _translate_openai(self, text):
        if not self.api_key:
            return text # Just return original if no key
        try:
            client = openai.OpenAI(api_key=self.api_key)
            
            # Constructing a prompt that includes the SAP glossary
            glossary_str = ", ".join([f"{k} -> {v}" for k, v in self.glossary.items()])
            
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": f"You are a professional SAP consultant and translator. Translate the following text from English to Korean. Keep SAP terminology as defined: {glossary_str}. If a term is not in the glossary but is technical SAP jargon, keep the English term in parentheses or follow standard SAP KR terminology. Maintain the tone of a professional business presentation. IMPORTANT: Output ONLY the translated Korean text."},
                    {"role": "user", "content": text}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"OpenAI Error: {e}")
            return text # Fallback to original text on error
