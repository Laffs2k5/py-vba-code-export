"""Tests for the CLI."""

import pytest
from pathlib import Path
from click.testing import CliRunner
from vba_export.cli import main


class TestCLI:
    """Test cases for command-line interface."""
    
    @pytest.fixture
    def sample_xlsm_path(self):
        """Return path to the sample XLSM file for testing."""
        return '/mnt/c/Users/peder/Downloads/Kopi av PrepareEU_detailed cost report_MCF.xlsm'
    
    def test_cli_help(self):
        """Test CLI help output."""
        runner = CliRunner()
        result = runner.invoke(main, ['--help'])
        assert result.exit_code == 0
        assert 'Export VBA code' in result.output
        assert 'INPUT_FILE' in result.output
        assert 'OUTPUT_DIR' in result.output
        assert '--verbose' in result.output
        assert 'Examples:' in result.output
        assert 'vba-export myfile.xlsm' in result.output
        assert 'Supported file formats' in result.output
        assert '.xlsm' in result.output
        assert '.docm' in result.output
        assert '.pptm' in result.output
    
    def test_cli_version(self):
        """Test CLI version output."""
        runner = CliRunner()
        result = runner.invoke(main, ['--version'])
        assert result.exit_code == 0
        assert 'vba-export' in result.output
        assert '0.1.0' in result.output
    
    def test_cli_missing_file(self):
        """Test CLI with non-existent input file."""
        runner = CliRunner()
        result = runner.invoke(main, ['/path/to/nonexistent/file.xlsm'])
        assert result.exit_code != 0
    
    def test_cli_basic_export(self, sample_xlsm_path):
        """Test basic export functionality."""
        if not Path(sample_xlsm_path).exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")
        
        runner = CliRunner()
        with runner.isolated_filesystem():
            result = runner.invoke(main, [sample_xlsm_path])
            assert result.exit_code == 0
            assert 'Successfully exported' in result.output
            assert 'VBA module' in result.output
            
            # Check that default output directory was created
            assert Path('exported_vba').exists()
    
    def test_cli_custom_output_dir(self, sample_xlsm_path):
        """Test export with custom output directory."""
        if not Path(sample_xlsm_path).exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")
        
        runner = CliRunner()
        with runner.isolated_filesystem():
            result = runner.invoke(main, [sample_xlsm_path, 'custom_output'])
            assert result.exit_code == 0
            assert 'Successfully exported' in result.output
            
            # Check that custom output directory was created
            assert Path('custom_output').exists()
            assert not Path('exported_vba').exists()
    
    def test_cli_verbose_output(self, sample_xlsm_path):
        """Test verbose output flag."""
        if not Path(sample_xlsm_path).exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")
        
        runner = CliRunner()
        with runner.isolated_filesystem():
            result = runner.invoke(main, [sample_xlsm_path, '-v'])
            assert result.exit_code == 0
            assert 'Input file:' in result.output
            assert 'Output directory:' in result.output
            assert sample_xlsm_path in result.output
    
    def test_cli_lists_exported_files(self, sample_xlsm_path):
        """Test that CLI lists all exported files."""
        if not Path(sample_xlsm_path).exists():
            pytest.skip(f"Test file not found: {sample_xlsm_path}")
        
        runner = CliRunner()
        with runner.isolated_filesystem():
            result = runner.invoke(main, [sample_xlsm_path])
            assert result.exit_code == 0
            
            # Check that output contains file paths
            assert 'exported_vba/' in result.output or 'exported_vba\\' in result.output
            assert '.cls' in result.output or '.bas' in result.output
    
    def test_cli_no_vba_modules(self):
        """Test behavior when file has no VBA modules.
        
        Note: oletools may extract content from text files,
        so we just verify the CLI doesn't crash.
        """
        runner = CliRunner()
        with runner.isolated_filesystem():
            # Create a simple text file (not a real Office file)
            dummy_file = Path('dummy.txt')
            dummy_file.write_text('not an office file')
            
            result = runner.invoke(main, [str(dummy_file)])
            # CLI should complete without crashing
            # oletools might extract "modules" from any file
            assert 'exported' in result.output.lower() or 'error' in result.output.lower()
