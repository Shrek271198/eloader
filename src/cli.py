
import click

ascii_snek = """\
    --..,_                     _,.--.
       `'.'.                .'`__ o  `;__.
          '.'.            .'.'`  '---'`  `
            '.`'--....--'`.'
              `'--....--'`
"""

def inputs_exist():

    return path.exists("MEL.xlsx") and path.exists("standard_details.xlsx")


@click.command()
@click.argument('standards.xlsx', type=click.Path(exists=True))
@click.argument('mel.xlsx', type=click.Path(exists=True))
def main():

    print(ascii_snek)

if __name__ == '__main__':

    main()
