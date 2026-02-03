"""Tests for the VBAExporter class."""

import pytest
from pathlib import Path
from vba_export.exporter import VBAExporter


class TestVBAExporter:
    """Test cases for VBAExporter."""

    @pytest.fixture
    def sample_xlsm_path(self):
        """Return path to the sample XLSM file for testing."""
        return Path('/mnt/c/Users/peder/Downloads/Kopi av PrepareEU_detailed cost report_MCF.xlsm')

    @pytest.fixture
    def temp_output_dir(self, tmp_path):
        """Create a temporary output directory."""
        return tmp_path / "test_export"

    def test_init(self, sample_xlsm_path):
        """Test exporter initialization."""
        exporter = VBAExporter(sample_xlsm_path)
        assert exporter.file_path == sample_xlsm_path

    def test_init_with_nonexistent_file(self):
        """Test initialization with non-existent file."""
        non_existent = Path("/path/to/nonexistent/file.xlsm")
        exporter = VBAExporter(non_existent)
        assert exporter.file_path == non_existent

    def test_get_extension_thisworkbook(self):
        """Test extension detection for ThisWorkbook module."""
        exporter = VBAExporter(Path("dummy.xlsm"))
        assert exporter._get_extension("ThisWorkbook") == ".cls"
        assert exporter._get_extension("thisworkbook") == ".cls"
        # If already has extension, don't add another
        assert exporter._get_extension("ThisWorkbook.cls") == ""

    def test_get_extension_sheet(self):
        """Test extension detection for Sheet modules."""
        exporter = VBAExporter(Path("dummy.xlsm"))
        assert exporter._get_extension("Sheet1") == ".cls"
        assert exporter._get_extension("sheet2") == ".cls"
        assert exporter._get_extension("SheetData") == ".cls"
        # If already has extension, don't add another
        assert exporter._get_extension("Sheet1.cls") == ""

    def test_get_extension_explicit(self):
        """Test extension detection for modules with explicit extensions."""
        exporter = VBAExporter(Path("dummy.xlsm"))
        # Modules that already have extensions should return empty string
        assert exporter._get_extension("Module1.bas") == ""
        assert exporter._get_extension("MyClass.cls") == ""
        assert exporter._get_extension("UserForm.frm") == ""

    def test_get_extension_default(self):
        """Test default extension for unknown module types."""
        exporter = VBAExporter(Path("dummy.xlsm"))
        assert exporter._get_extension("SomeModule") == ".bas"
        assert exporter._get_extension("CustomCode") == ".bas"

    def test_extract_vba_modules(self, sample_xlsm_path):
        """Test VBA module extraction from real file."""
        if not sample_xlsm_path.exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")

        exporter = VBAExporter(sample_xlsm_path)
        modules = exporter.extract_vba_modules()

        assert isinstance(modules, dict)
        assert len(modules) > 0

        # Check that we have module names and code
        for module_name, code in modules.items():
            assert isinstance(module_name, str)
            assert len(module_name) > 0
            assert isinstance(code, str)
            assert len(code) > 0

    def test_export_to_directory(self, sample_xlsm_path, temp_output_dir):
        """Test exporting VBA modules to directory."""
        if not sample_xlsm_path.exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")

        exporter = VBAExporter(sample_xlsm_path)
        created_files = exporter.export_to_directory(temp_output_dir)

        # Check that files were created
        assert len(created_files) > 0
        assert temp_output_dir.exists()

        # Check that all files exist and have content
        for file_path in created_files:
            assert file_path.exists()
            assert file_path.is_file()
            assert file_path.stat().st_size > 0

            # Verify file has correct extension
            assert file_path.suffix in ['.bas', '.cls', '.frm']

            # Verify file content is readable
            content = file_path.read_text(encoding='utf-8', errors='ignore')
            assert len(content) > 0

    def test_export_creates_directory(self, sample_xlsm_path, temp_output_dir):
        """Test that export creates output directory if it doesn't exist."""
        if not sample_xlsm_path.exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")

        assert not temp_output_dir.exists()

        exporter = VBAExporter(sample_xlsm_path)
        exporter.export_to_directory(temp_output_dir)

        assert temp_output_dir.exists()
        assert temp_output_dir.is_dir()

    def test_export_file_naming(self, sample_xlsm_path, temp_output_dir):
        """Test that exported files have correct names and extensions."""
        if not sample_xlsm_path.exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")

        exporter = VBAExporter(sample_xlsm_path)
        created_files = exporter.export_to_directory(temp_output_dir)

        for file_path in created_files:
            # File should have single extension only
            assert file_path.suffix in ['.bas', '.cls', '.frm']
            # Verify no double extensions
            name_without_ext = file_path.stem
            assert not name_without_ext.endswith('.bas')
            assert not name_without_ext.endswith('.cls')
            assert not name_without_ext.endswith('.frm')
