"""Core VBA code extraction logic."""

from pathlib import Path
from typing import List, Dict
from oletools.olevba import VBA_Parser


class VBAExporter:
    """Extract VBA code from macro-enabled Office files."""

    def __init__(self, file_path: Path):
        """Initialize the exporter with a file path.

        Args:
            file_path: Path to the macro-enabled Office file
        """
        self.file_path = file_path

    def _get_extension(self, vba_filename: str) -> str:
        """Determine file extension based on module name and type.

        Args:
            vba_filename: Name of the VBA module

        Returns:
            File extension including the dot (e.g., '.bas', '.cls', '.frm'),
            or empty string if the filename already has a valid extension
        """
        name = vba_filename.lower()

        # Check if the filename already has a VBA extension
        if name.endswith(".bas") or name.endswith(".cls") or name.endswith(".frm"):
            return ""

        # Determine extension based on module type
        if name.startswith("thisworkbook"):
            return ".cls"
        elif name.startswith("sheet"):
            return ".cls"
        else:
            # Default to .bas if unknown
            return ".bas"

    def extract_vba_modules(self) -> Dict[str, str]:
        """Extract all VBA modules from the file.

        Returns:
            Dictionary mapping module names to their source code
        """
        modules = {}
        vba = VBA_Parser(str(self.file_path))

        try:
            for (filename, stream_path, vba_filename, vba_code) in vba.extract_all_macros():
                if vba_filename is None:
                    continue
                modules[vba_filename] = vba_code
        finally:
            vba.close()

        return modules

    def export_to_directory(self, output_dir: Path) -> List[Path]:
        """Export VBA modules to separate files in a directory.

        Args:
            output_dir: Directory to save the exported VBA files

        Returns:
            List of paths to the created files
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        created_files = []
        modules = self.extract_vba_modules()

        for vba_filename, vba_code in modules.items():
            extension = self._get_extension(vba_filename)
            out_path = output_dir / (vba_filename + extension)

            with open(out_path, "w", encoding="utf-8", errors="ignore") as f:
                f.write(vba_code)

            created_files.append(out_path)

        return created_files
