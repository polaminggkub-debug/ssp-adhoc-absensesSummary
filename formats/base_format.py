"""
Base class for format handlers.

Each format handler extracts employee absence data from a specific Excel format.
"""

from abc import ABC, abstractmethod
from typing import List, Dict, Any


class BaseFormatHandler(ABC):
    """Abstract base class for Excel format handlers."""

    @abstractmethod
    def extract_employees(self, filepath: str) -> List[Dict[str, Any]]:
        """
        Extract employee data from an Excel file.

        Args:
            filepath: Path to the Excel file

        Returns:
            List of employee dictionaries with keys:
                - primary_key: Matching key for deduplication
                - name_key: Name-based key
                - emp_id: Employee ID string
                - nickname: Thai nickname (if any)
                - display_name: Formatted display name
                - note: Notes from name field (quit dates, etc.)
                - position: Job position
                - department: Department name
                - payType: Payment type
                - totals: List of 17 absence type totals
        """
        pass

    @abstractmethod
    def get_format_config(self) -> Dict[str, Any]:
        """
        Return the format configuration dictionary.

        Returns:
            Configuration dict with column mappings and settings
        """
        pass

    @property
    @abstractmethod
    def format_name(self) -> str:
        """Return human-readable format name for logging."""
        pass
