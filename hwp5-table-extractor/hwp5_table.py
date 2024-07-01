import re
import struct
from io import BytesIO
from itertools import islice

import olefile

from enums import tag_table, control_char_table
from utils import ZlibDecompressStream

class Record(object):
    def __init__(self, tag_id, payload, level,  parent=None):
        self.parent = parent
        self.children = []

        self.tag_id = tag_id
        self.tag_name = tag_table.get(self.tag_id, '<ROOT>')
        self.level = level
        self.payload = payload

    def __repr__(self):
        return '<Record %s>' % self.tag_name

    def get_next_siblings(self, count=None):
        start_idx = self.parent.children.index(self) + 1
        if count is None:
            end_idx = None
        else:
            end_idx = start_idx + count

        return islice(self.parent.children, start_idx, end_idx)

    @staticmethod
    def build_tree_from_stream(stream):
        root = Record(None, None, None)

        while True:
            header = stream.read(4)
            if not header:
                break

            header = struct.unpack('<I', header)[0]

            tag_id = header & 0x3ff
            level = (header >> 10) & 0x3ff
            size = (header >> 20) & 0xfff

            if size == 0xfff:
                size = struct.unpack('<I', stream.read(4))[0]

            payload = stream.read(size)

            last_record = root
            for _ in range(level):
                last_record = last_record.children[-1]

            last_record.children.append(Record(tag_id, payload, level, last_record))

        return root

    def get_text(self):
        regex = re.compile(rb'([\x00-\x1f])\x00')

        text = ''

        cursor_idx = 0
        search_idx = 0

        while cursor_idx < len(self.payload):
            if search_idx < cursor_idx:
                search_idx = cursor_idx

            searched = regex.search(self.payload, search_idx)
            if searched:
                pos = searched.start()

                if pos & 1:
                    search_idx = pos + 1
                elif pos > cursor_idx:
                    text += self.payload[cursor_idx:pos].decode('utf-16')
                    cursor_idx = pos
                else:
                    control_char = ord(searched.group(1))
                    control_char_size = control_char_table[control_char][1].size

                    if control_char == 0x0a:
                        text += '\n'

                    cursor_idx = pos + control_char_size * 2
            else:
                text += self.payload[search_idx:].decode('utf-16')
                break

        return text

class Table(object):
    def __init__(self, caption, row_cnt, col_cnt):
        self.caption = caption
        self.row_cnt = row_cnt
        self.col_cnt = col_cnt

        self.rows = [[] for _ in range(row_cnt)]

    def __repr__(self):
        return '<Table %s>' % self.caption

class TableCell(object):
    def __init__(self, lines, row, col, row_span, col_span):
        self.lines = lines
        self.row = row
        self.col = col
        self.row_span = row_span
        self.col_span = col_span

    def __repr__(self):
        return '<TableCell(%d, %d) %s>' % (self.row, self.col, self.lines)

class TableCell(object):
    def __init__(self, lines, row, col, row_span, col_span):
        self.lines = lines
        self.row = row
        self.col = col
        self.row_span = row_span
        self.col_span = col_span

    def __repr__(self):
        return '<TableCell(%d, %d) %s>' % (self.row, self.col, self.lines)

def make_tables(record_tree_root):
    def traverse(record, depth=0):
        # print('  ' * depth + repr(record))
        # if (record.tag_name == 'HWPTAG_PARA_TEXT'
        #     and record.parent.parent.tag_name == '<ROOT>'
        #     and record.payload[0] != 0x0b):
        #     ctx['table_caption'] = record.get_text().strip()
        if record.tag_name == 'HWPTAG_TABLE':
            if 'current_table_idx' not in ctx:
                ctx['current_table_idx'] = 0
            else:
                ctx['current_table_idx'] += 1

            row_cnt = struct.unpack('<H', record.payload[4:6])[0]
            col_cnt = struct.unpack('<H', record.payload[6:8])[0]

            ctx['tables'].append(Table(None, row_cnt, col_cnt))
            # ctx['tables'].append(Table(ctx['table_caption'], row_cnt, col_cnt))
        elif (record.tag_name == 'HWPTAG_LIST_HEADER'
              and record.parent.tag_name == 'HWPTAG_CTRL_HEADER'
              and record.parent.payload[:4][::-1] == b'tbl '):
            paragraph_count = struct.unpack('<H', record.payload[:2])[0]
            col = struct.unpack('<H', record.payload[8:10])[0]
            row = struct.unpack('<H', record.payload[10:12])[0]
            col_span = struct.unpack('<H', record.payload[12:14])[0]
            row_span = struct.unpack('<H', record.payload[14:16])[0]

            lines = []
            for sibling in record.get_next_siblings(paragraph_count):
                for child in sibling.children:
                    if child.tag_name == 'HWPTAG_PARA_TEXT':
                        lines.extend(child.get_text().strip().splitlines())
                        break
            #print("len(ctx['tables'])", len(ctx['tables']))
            #print("ctx['current_table_idx']", ctx['current_table_idx'])
            #print("row", row)
            if ctx['current_table_idx'] < len(ctx['tables']) and row < len(ctx['tables'][ctx['current_table_idx']].rows):
                ctx['tables'][ctx['current_table_idx']].rows[row].append(
                    TableCell(lines, row, col, row_span, col_span)
                )
            else:
                print(f"Invalid table index or row index, row_num[{row }], skipping append")

        for child in record.children:
            traverse(child, depth + 1)

    ctx = {'tables': []}
    traverse(record_tree_root)
    #print('ctx:', ctx)
    #print('re#################')

    #return ctx['tables'] # 원래는 리스트로 리턴하는거였음
    result = {idx: table for idx, table in enumerate(ctx['tables'])}

    return result


def make_paragraph(record_tree_root, tag_table_position=True):
    idx = 0
    table_idx = 0

    def traverse(record, tag_table_position, depth=0, ):
        nonlocal idx  # idx를 외부 변수로 사용하겠다고 명시
        nonlocal table_idx

        paragraphs = {}

        if record.tag_name == 'HWPTAG_PARA_TEXT' and record.level == 1: #(level3이면 테이블 내 텍스트도 함께 표시됨)
            paragraphs[idx] = record.get_text().strip()
            idx += 1  # idx를 증가시킵니다.
        elif record.tag_name == 'HWPTAG_TABLE':
            if tag_table_position:
                paragraphs[idx] = f"[[{table_idx}]]"
            idx +=1
            table_idx+=1

        for child in record.children:
            child_paragraphs = traverse(child, depth + 1)
            paragraphs.update(child_paragraphs)

        return paragraphs

    result = traverse(record_tree_root, tag_table_position)
    print('all:', result)
    return result

class HwpFile(object):
    def __init__(self, file):
        self.ole = olefile.OleFileIO(file)

    @property
    def compressed(self):
        if not hasattr(self, '_compressed'):
            with self.ole.openstream('FileHeader') as stream:
                stream.seek(36)
                flag = struct.unpack('<I', stream.read(4))[0]
                self._compressed = bool(flag & 1)

        return self._compressed

    def get_body_stream(self, section_idx):
        if not self.ole.exists('BodyText/Section%d' % section_idx):
            raise IndexError('Section%d does not exist' % section_idx)

        return self.ole.openstream('BodyText/Section%d' % section_idx)

    def get_record_tree(self, section_idx):
        with self.get_body_stream(section_idx) as stream:
            if self.compressed:
                stream = ZlibDecompressStream(stream, -15)

            record_tree_root = Record.build_tree_from_stream(stream)

        return record_tree_root

    def get_tables_by_list(self, section_idx):
        tables = []
        for k, v in make_tables(self.get_record_tree(section_idx)).items():
            tables.append(v)
        return tables
    def get_tables(self, section_idx):
        ## 번호가 키인 딕셔너리 형태로 리턴
        return make_tables(self.get_record_tree(section_idx))
    
    
    def get_paragraphs(self, section_idx):
        record_tree = self.get_record_tree(section_idx)
        return make_paragraph(record_tree)