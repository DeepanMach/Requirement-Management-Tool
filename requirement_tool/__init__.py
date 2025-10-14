"""Core package for the requirement management tool."""

from .data_manager import RequirementDataManager, RequirementDataError

__all__ = [
    "RequirementDataManager",
    "RequirementDataError",
]
