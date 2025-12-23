"""JSON schema definitions and helpers"""
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any
import json


@dataclass
class Resident:
    name: str
    rotation_type: str
    ng_dates: List[str] = field(default_factory=list)
    ng_reasons: Dict[str, List[str]] = field(default_factory=dict)
    source_rows: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class ShiftJSON:
    month: str
    residents: List[Resident] = field(default_factory=list)
    unknown_names: List[str] = field(default_factory=list)
    parse_errors: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self, indent=2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)
