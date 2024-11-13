from copy import deepcopy

from docx import Document

from read_excel import get_table


def _replase_paragraph_text(key: str, text: str, paragraph: object) -> object:
    if key in paragraph.text:
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(str(key), str(text))
    return paragraph


def create_word_file(table: str, template: str, path_: str) -> None:
    docs_file = ""
    base_content = Document(template)
    table = get_table(table)

    for docx in table:
        content = deepcopy(base_content)
        for k, v in docx.items():
            for paragraph in content.paragraphs:
                paragraph = _replase_paragraph_text(k, v, paragraph)

        docs_file = str(docx.get("index"))
        content.save(f"{path_}\\{docs_file}.docx")


if __name__ == "__main__":
    table = ".\\tests\\data\\test_table.xlsx"
    template = ".\\tests\\data\\test_template.docx"
    path_ = ".\\tests\\data\\"
    create_word_file(table=table, template=template, path_=path_)
