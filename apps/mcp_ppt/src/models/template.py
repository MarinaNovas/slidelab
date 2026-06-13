# src/models/template.py

from dataclasses import dataclass


@dataclass
class SemanticTemplate:
    profile: dict

    def get_layout_info(self, semantic_type: str) -> dict:
        if semantic_type not in self.profile:
            raise ValueError(f"Layout semantic type '{semantic_type}' not found")

        return self.profile[semantic_type]