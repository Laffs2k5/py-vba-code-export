"""Command-line interface for VBA code exporter."""

import click
from pathlib import Path
from vba_export.exporter import VBAExporter


@click.command(
    epilog="""
Examples:

  Export VBA from an Excel file to default directory (exported_vba/):

    \b
    $ vba-export myfile.xlsm

  Export to a custom directory:

    \b
    $ vba-export myfile.xlsm output/vba_code

  Export with verbose output:

    \b
    $ vba-export myfile.xlsm -v

Supported file formats: .xlsm (Excel), .docm (Word), .pptm (PowerPoint)
"""
)
@click.argument(
    'input_file',
    type=click.Path(exists=True, path_type=Path),
    metavar='INPUT_FILE'
)
@click.argument(
    'output_dir',
    type=click.Path(path_type=Path),
    required=False,
    metavar='[OUTPUT_DIR]'
)
@click.option(
    '--verbose', '-v',
    is_flag=True,
    help='Show detailed progress information'
)
@click.version_option(
    version='0.1.0',
    prog_name='vba-export',
    message='%(prog)s version %(version)s'
)
def main(input_file: Path, output_dir: Path, verbose: bool):
    """Export VBA code from macro-enabled Microsoft Office files.
    
    This tool extracts VBA macros from Office files in OOXML format (.xlsm, .docm, 
    .pptm) and saves them as separate text files (.bas, .cls, .frm) that can be 
    imported back into Office applications or used for version control.
    
    \b
    Arguments:
      INPUT_FILE   Path to the macro-enabled Office file
      OUTPUT_DIR   Directory for exported files (default: exported_vba/)
    """
    if output_dir is None:
        output_dir = Path('exported_vba')
    
    if verbose:
        click.echo(f"Input file: {input_file}")
        click.echo(f"Output directory: {output_dir}")
    
    try:
        exporter = VBAExporter(input_file)
        files_created = exporter.export_to_directory(output_dir)
        
        if files_created:
            click.echo(f"Successfully exported {len(files_created)} VBA module(s):")
            for file_path in files_created:
                click.echo(f"  - {file_path}")
        else:
            click.echo("No VBA modules found in the file.")
    
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        raise click.Exit(1)
