from click.testing import CliRunner
from cli import main


def test_input_arguments():
    runner = CliRunner()
    result = runner.invoke(main)
    assert result.exit_code == 2
