
import click

@click.command()
@click.argument('standards', type=click.Path(exists=True))
@click.argument('mel', type=click.Path(exists=True))
def main(standards, mel):
    """eMax ELoader Command Line Interface

    STANDARDS is an excel file that contains the project standard details.

    MEL is an excel file that contains the Mechanical Equipment List.
    """
    print(standards)
    print(mel)


if __name__ == '__main__':

    main()
