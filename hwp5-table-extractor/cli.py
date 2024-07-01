import click
from jinja2 import Template

from hwp5_table import HwpFile

@click.command()
@click.argument('input', type=click.File('rb'))
# @click.argument('output', type=click.File('w'))
@click.argument('output', type=click.File('w', encoding='utf-8'))
def cli(input, output):
    hwp = HwpFile(input)

    tables = {}
    section_idx = 0
    while hwp.ole.exists('BodyText/Section%d' % section_idx):
        #tables.extend(hwp.get_tables_by_list(section_idx))
        tables = hwp.get_tables(section_idx)
        #tables.extend(hwp.get_tables(section_idx))
        print('#######################')
        hwp.get_paragraphs(section_idx)
        print('#######################')
        section_idx += 1

    # Render to html using Jinja2 template engine
    template = Template('''
      <!doctype html>
      <html>
      <head>
        <style type="text/css">
          body {
            padding: 20px;
          }
          table {
            width: 100%;
            max-width: 1000px;
            border-collapse: collapse;
            margin: 0 auto;
            margin-bottom: 30px;
          }
          td {
            font-size: 13px;
            border: 1px solid #aaa;
            padding: 5px;
            text-align: center;
          }
        </style>
      </head>
      <body>
        {% for k, table in tables.items() %}
        <table>
          <p>TABLE_IDX: {{k}}</p>
          <tbody>
          {% for row in table.rows %}
            <tr>
            {% for cell in row %}
              <td rowspan="{{ cell.row_span }}" colspan="{{ cell.col_span }}">{{ '<br>'.join(cell.lines) }}</td>
            {% endfor %}
            </tr>
          {% endfor %}
          </tbody>
        </table>
        {% endfor %}
      </body>
      </html>
    ''')

    output.write(template.render(tables=tables))

if __name__ == '__main__':
    cli()