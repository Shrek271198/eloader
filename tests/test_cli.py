from click.testing import CliRunner
from cli import main


def test_missing_input_arguments():
    runner = CliRunner()
    result = runner.invoke(main)
    assert result.exit_code == 2

def test_available_input_arguments():
    runner = CliRunner()
    result = runner.invoke(main, ['tests/fixtures/standards.xlsx','tests/fixtures/mel.xlsx'])
    assert result.exit_code == 0
