import xml.etree.ElementTree as ET

class XMLSheet:
    def __init__(self, name, headers, rows):
        self.name = name
        self.headers = headers  # List[str]
        self.rows = rows        # List[List[str]]

class XMLParser:
    @staticmethod
    def parse_file(file_path):
        """
        Парсит XML-файл и возвращает список объектов XMLSheet.
        """
        with open(file_path, 'rb') as f:
            raw_bytes = f.read()

        root = ET.fromstring(raw_bytes)

        # 1. Проверяем, является ли документ Excel XML Spreadsheet
        tag_lower = root.tag.lower()
        if tag_lower.endswith('workbook') or 'spreadsheet' in root.tag.lower():
            sheets = XMLParser._parse_excel_xml(root)
            if sheets:
                return sheets

        # 2. Проверяем обобщенные табличные XML с повторяющимися элементами
        sheets = XMLParser._parse_generic_tabular_xml(root)
        if sheets:
            return sheets

        # 3. Резервный вариант: Плоское представление дерева XML
        return [XMLParser._parse_flattened_xml(root)]

    @staticmethod
    def _clean_number_string(val):
        """
        Нормализует некрасивое представление чисел в XML (например "4." -> "4", ".35" -> "0.35").
        """
        if not val:
            return val
        
        # Исправление лидирующей точки: ".35" -> "0.35", "-.35" -> "-0.35"
        if val.startswith('.'):
            val = '0' + val
        elif val.startswith('-.'):
            val = '-0' + val[1:]

        # Исправление висячей точки на конце: "4." -> "4"
        if val.endswith('.'):
            val = val[:-1]

        return val

    @staticmethod
    def _parse_excel_xml(root):
        sheets = []

        worksheets = [elem for elem in root.iter() if elem.tag.lower().endswith('worksheet')]

        for ws in worksheets:
            name = "Лист"
            for attr_k, attr_v in ws.attrib.items():
                if attr_k.endswith('Name') or attr_k.endswith('name'):
                    name = attr_v
                    break

            table = next((elem for elem in ws.iter() if elem.tag.lower().endswith('table')), None)
            if table is None:
                continue

            rows_elems = [elem for elem in table if elem.tag.lower().endswith('row')]

            parsed_rows = []
            max_cols = 0

            for r_elem in rows_elems:
                row_cells = []
                col_idx = 1
                cell_elems = [elem for elem in r_elem if elem.tag.lower().endswith('cell')]

                for c_elem in cell_elems:
                    idx_val = None
                    for ak, av in c_elem.attrib.items():
                        if ak.endswith('Index') or ak.endswith('index'):
                            try:
                                idx_val = int(av)
                            except ValueError:
                                pass
                            break
                    if idx_val:
                        col_idx = idx_val

                    data_elem = next((e for e in c_elem if e.tag.lower().endswith('data')), None)
                    
                    val = ""
                    if data_elem is not None and data_elem.text:
                        val = data_elem.text.strip()
                    elif c_elem.text and c_elem.text.strip():
                        val = c_elem.text.strip()

                    # Очистка чисел от лидирующих/висячих точек
                    val = XMLParser._clean_number_string(val)

                    while len(row_cells) < col_idx - 1:
                        row_cells.append("")
                    row_cells.append(val)
                    col_idx += 1

                if any(row_cells):
                    max_cols = max(max_cols, len(row_cells))
                    parsed_rows.append(row_cells)

            if not parsed_rows:
                continue

            for row in parsed_rows:
                while len(row) < max_cols:
                    row.append("")

            headers = [f"Колонка {i+1}" for i in range(max_cols)]
            header_row_idx = -1

            for idx, row in enumerate(parsed_rows):
                non_empty = [c for c in row if c]
                if len(non_empty) >= 1:
                    headers = [c if c else f"Колонка {i+1}" for i, c in enumerate(row)]
                    header_row_idx = idx
                    break

            data_rows = parsed_rows[header_row_idx + 1:] if header_row_idx != -1 else parsed_rows
            sheets.append(XMLSheet(name=name, headers=headers, rows=data_rows))

        return sheets

    @staticmethod
    def _parse_generic_tabular_xml(root):
        candidates = []
        for elem in root.iter():
            children = list(elem)
            if not children:
                continue
            
            tag_counts = {}
            for child in children:
                tag_counts[child.tag] = tag_counts.get(child.tag, 0) + 1

            for tag, count in tag_counts.items():
                if count >= 2:
                    candidates.append((elem, tag, count))

        if not candidates:
            return None

        candidates.sort(key=lambda x: x[2], reverse=True)
        container, row_tag, count = candidates[0]

        row_elems = [e for e in container if e.tag == row_tag]
        
        headers = []
        for r_elem in row_elems:
            for child in r_elem:
                clean_tag = XMLParser._clean_tag(child.tag)
                if clean_tag not in headers:
                    headers.append(clean_tag)
                for attr_k in child.attrib:
                    attr_col = f"{clean_tag}@{attr_k}"
                    if attr_col not in headers:
                        headers.append(attr_col)

        if not headers:
            return None

        rows = []
        for r_elem in row_elems:
            row_dict = {}
            for child in r_elem:
                clean_tag = XMLParser._clean_tag(child.tag)
                raw_txt = child.text.strip() if child.text else ""
                row_dict[clean_tag] = XMLParser._clean_number_string(raw_txt)
                for attr_k, attr_v in child.attrib.items():
                    row_dict[f"{clean_tag}@{attr_k}"] = attr_v

            rows.append([row_dict.get(h, "") for h in headers])

        sheet_name = XMLParser._clean_tag(row_tag)
        return [XMLSheet(name=sheet_name, headers=headers, rows=rows)]

    @staticmethod
    def _parse_flattened_xml(root):
        headers = ["Индекс", "Тэг", "Путь", "Атрибуты", "Значение"]
        rows = []

        idx = 1
        def _traverse(node, current_path):
            nonlocal idx
            clean_t = XMLParser._clean_tag(node.tag)
            path = f"{current_path}/{clean_t}"
            attrs = ", ".join([f"{k}='{v}'" for k, v in node.attrib.items()])
            text = node.text.strip() if node.text and node.text.strip() else ""
            text = XMLParser._clean_number_string(text)

            if text or attrs:
                rows.append([str(idx), clean_t, path, attrs, text])
                idx += 1

            for child in node:
                _traverse(child, path)

        _traverse(root, "")
        return XMLSheet(name="Иерархия XML", headers=headers, rows=rows)

    @staticmethod
    def _clean_tag(tag):
        if '}' in tag:
            return tag.split('}', 1)[1]
        return tag
