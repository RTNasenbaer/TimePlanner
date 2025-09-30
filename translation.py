#!/usr/bin/env python3
"""
Translation system for TimePlanner
Provides easy-to-use translation functionality for multiple languages
"""

import json
import os
from typing import Dict, Optional


class TranslationManager:
    """Manages translations for the TimePlanner application"""
    
    def __init__(self, language: str = "de-de", fallback_language: str = "de-de"):
        """
        Initialize the translation manager
        
        Args:
            language: Primary language code (e.g., "en-us", "de-de")
            fallback_language: Fallback language if primary is not available
        """
        self.language = language
        self.fallback_language = fallback_language
        self.translations: Dict[str, str] = {}
        self.fallback_translations: Dict[str, str] = {}
        self.lang_dir = os.path.join(os.path.dirname(__file__), "lang")
        
        self.load_translations()
    
    def load_translations(self):
        """Load translation files"""
        # Load primary language
        self._load_language_file(self.language, self.translations)
        
        # Load fallback language if different
        if self.language != self.fallback_language:
            self._load_language_file(self.fallback_language, self.fallback_translations)
    
    def _load_language_file(self, language: str, target_dict: Dict[str, str]):
        """Load a specific language file"""
        lang_file = os.path.join(self.lang_dir, f"{language}.json")
        
        if os.path.exists(lang_file):
            try:
                with open(lang_file, 'r', encoding='utf-8') as f:
                    target_dict.update(json.load(f))
                print(f"Loaded {len(target_dict)} translations for {language}")
            except (json.JSONDecodeError, FileNotFoundError) as e:
                print(f"Error loading {language} translations: {e}")
        else:
            print(f"Translation file not found: {lang_file}")
    
    def tr(self, key: str, **kwargs) -> str:
        """
        Translate a key with optional formatting arguments
        
        Args:
            key: Translation key
            **kwargs: Formatting arguments for string interpolation
            
        Returns:
            Translated and formatted string
        """
        # Try primary language first
        text = self.translations.get(key)
        
        # Fall back to fallback language
        if text is None and self.fallback_translations:
            text = self.fallback_translations.get(key)
        
        # If still not found, return the key itself as fallback
        if text is None:
            text = key
            print(f"Translation missing: {key}")
        
        # Apply formatting if arguments provided
        if kwargs:
            try:
                text = text.format(**kwargs)
            except (KeyError, ValueError) as e:
                print(f"Translation formatting error for '{key}': {e}")
        
        return text
    
    def change_language(self, language: str):
        """
        Change the active language
        
        Args:
            language: New language code
        """
        if language != self.language:
            self.language = language
            self.translations.clear()
            self.load_translations()
    
    def get_available_languages(self) -> list:
        """Get list of available language codes"""
        if not os.path.exists(self.lang_dir):
            return []
        
        languages = []
        for file in os.listdir(self.lang_dir):
            if file.endswith('.json'):
                lang_code = file[:-5]  # Remove .json extension
                languages.append(lang_code)
        
        return sorted(languages)
    
    def get_language_display_name(self, language_code: str) -> str:
        """Get display name for a language code"""
        display_names = {
            "de-de": "Deutsch (Deutschland)",
            "en-us": "English (United States)",
            "en-gb": "English (United Kingdom)",
            "fr-fr": "Français (France)",
            "es-es": "Español (España)",
            "it-it": "Italiano (Italia)",
            "nl-nl": "Nederlands (Nederland)",
            "pt-pt": "Português (Portugal)",
        }
        return display_names.get(language_code, language_code.upper())


# Global translation manager instance
# This will be initialized when first imported
_translation_manager: Optional[TranslationManager] = None


def init_translation_system(language: str = "de-de") -> TranslationManager:
    """
    Initialize the global translation system
    
    Args:
        language: Initial language code
        
    Returns:
        TranslationManager instance
    """
    global _translation_manager
    _translation_manager = TranslationManager(language)
    return _translation_manager


def tr(key: str, **kwargs) -> str:
    """
    Global translation function - shorthand for easy use
    
    Args:
        key: Translation key
        **kwargs: Formatting arguments
        
    Returns:
        Translated string
    """
    global _translation_manager
    if _translation_manager is None:
        _translation_manager = TranslationManager()
    
    return _translation_manager.tr(key, **kwargs)


def get_translation_manager() -> TranslationManager:
    """Get the global translation manager instance"""
    global _translation_manager
    if _translation_manager is None:
        _translation_manager = TranslationManager()
    
    return _translation_manager


def change_language(language: str):
    """Change the global language setting"""
    global _translation_manager
    if _translation_manager is None:
        _translation_manager = TranslationManager(language)
    else:
        _translation_manager.change_language(language)


# Usage Examples:
"""
# Basic usage:
from translation import tr

# Simple translation
title = tr("app_title")  # "Zeitplaner"

# Translation with formatting
duration_info = tr("info_duration", hours=2, minutes=30, total=150)
# "Gesamtdauer: 2h 30min (150 Minuten)"

# Error message with parameter
error_msg = tr("error_template_not_found", template="myfile.docx")
# "Template-Datei 'myfile.docx' nicht gefunden!"

# Change language
change_language("en-us")
title = tr("app_title")  # "Time Planner"

# Get available languages
tm = get_translation_manager()
languages = tm.get_available_languages()  # ["de-de", "en-us"]
"""