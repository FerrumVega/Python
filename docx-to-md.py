import re
import docx
import os

# === КОД БЫЛ СГЕНЕРИРОВАН НЕЙРОСЕТЬЮ ===


def sanitize_anchor(text):
    anchor = re.sub(r"[^\w\s-]", "", text.lower())
    anchor = re.sub(r"[\s-]+", "-", anchor)
    return anchor.strip("-")


def is_list_paragraph(paragraph):
    """Проверка, является ли параграф элементом списка в Word."""
    p = paragraph._p  # получаем xml-элемент параграфа
    numPr = p.find(".//w:numPr", p.nsmap)
    return numPr is not None


def add_two_spaces(line):
    """Добавляет два пробела в конец строки."""
    return line + "  "


def docx_to_markdown(docx_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    doc = docx.Document(docx_path)

    markdown_lines = []
    current_code_block = []
    in_code_block = False
    image_counter = 1

    image_parts = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_parts[rel.rId] = rel.target_part

    for element in doc.element.body:
        if element.tag.endswith("p"):
            paragraph = docx.text.paragraph.Paragraph(element, doc)

            # Обработка изображений
            if paragraph._element.xpath(".//pic:pic"):
                for run in paragraph.runs:
                    if run._element.xpath(".//pic:pic"):
                        blip = run._element.xpath(".//a:blip/@r:embed")[0]
                        if blip in image_parts:
                            img_path = os.path.join(
                                images_dir, f"image_{image_counter}.png"
                            )
                            with open(img_path, "wb") as f:
                                f.write(image_parts[blip].blob)
                            markdown_lines.append(
                                add_two_spaces(
                                    f"![Image {image_counter}](images/image_{image_counter}.png)"
                                )
                            )
                            image_counter += 1
                continue  # Пропускаем дальнейшую обработку для параграфов с изображениями

            is_code_font = all(
                run.font and run.font.name == "Cascadia Mono" for run in paragraph.runs
            )
            has_content = bool(paragraph.text.strip())

            # Обработка пустых параграфов
            if not has_content:
                if in_code_block and is_code_font:
                    current_code_block.append("")  # Пустая строка в блоке кода
                else:
                    markdown_lines.append("")  # Добавляем пустую строку в markdown
                continue

            if is_code_font:
                if not in_code_block and current_code_block:
                    markdown_lines.append(
                        add_two_spaces(
                            f"""```python
{'\n'.join(current_code_block)}
```"""
                        )
                    )
                    current_code_block = []
                in_code_block = True
                current_code_block.append(add_two_spaces(paragraph.text))
            else:
                if in_code_block and current_code_block:
                    markdown_lines.append(
                        add_two_spaces(
                            f"```python\n{'\n'.join(current_code_block)}\n```"
                        )
                    )
                    current_code_block = []
                    in_code_block = False

                paragraph_text = ""
                i = 0
                runs = paragraph.runs

                # Новая логика обработки инлайн-кода
                while i < len(runs):
                    run = runs[i]
                    if (
                        run.font
                        and run.font.name == "Cascadia Mono"
                        and run.text.strip()
                    ):
                        # Начало инлайн-кода
                        code_parts = [run.text]
                        j = i + 1
                        # Ищем следующие runs, которые могут быть частью этого же кода
                        while j < len(runs):
                            next_run = runs[j]
                            if (
                                next_run.font
                                and next_run.font.name == "Cascadia Mono"
                                or (
                                    next_run.text == " "
                                    and j + 1 < len(runs)
                                    and runs[j + 1].font
                                    and runs[j + 1].font.name == "Cascadia Mono"
                                )
                            ):
                                code_parts.append(next_run.text)
                                j += 1
                            else:
                                break
                        paragraph_text += f"`{''.join(code_parts)}`"
                        i = j
                    else:
                        paragraph_text += run.text
                        i += 1

                paragraph_text = add_two_spaces(paragraph_text.rstrip())

                if paragraph.style.name.startswith("Heading"):
                    level = int(paragraph.style.name.split()[-1])
                    clean_text = paragraph_text.strip()
                    markdown_lines.append(add_two_spaces(f"{'#' * level} {clean_text}"))
                else:
                    if is_list_paragraph(paragraph):
                        markdown_lines.append(
                            add_two_spaces(f"- {paragraph_text.strip()}")
                        )
                    else:
                        markdown_lines.append(add_two_spaces(paragraph_text))

        elif element.tag.endswith("tbl"):
            table = docx.table.Table(element, doc)
            markdown_lines.append(add_two_spaces("\n"))

            num_cols = len(table.columns)

            header = []
            for cell in table.rows[0].cells:
                header.append(add_two_spaces(cell.text.strip()))
            markdown_lines.append(add_two_spaces("| " + " | ".join(header) + " |"))

            markdown_lines.append(
                add_two_spaces("|" + "|".join(["---"] * num_cols) + "|")
            )

            for row in table.rows[1:]:
                row_data = []
                for cell in row.cells:
                    row_data.append(
                        add_two_spaces(cell.text.strip().replace("\n", "<br>"))
                    )
                markdown_lines.append(
                    add_two_spaces("| " + " | ".join(row_data) + " |")
                )

            markdown_lines.append(add_two_spaces("\n"))

    if current_code_block:
        markdown_lines.append(
            add_two_spaces(f"```python\n{'\n'.join(current_code_block)}\n```")
        )

    # Добавляем два пробела в конец каждой строки в оглавлении
    toc = [add_two_spaces("## Оглавление")]
    headers = []

    for line in markdown_lines:
        if line.startswith("#") and not line.startswith("## Оглавление"):
            level = len(line.split()[0])
            title = " ".join(line.split()[1:])
            headers.append((level, title))

    anchor_counts = {}
    for level, title in headers:
        anchor = sanitize_anchor(title)
        if anchor in anchor_counts:
            anchor_counts[anchor] += 1
            anchor = f"{anchor}-{anchor_counts[anchor]}"
        else:
            anchor_counts[anchor] = 1

        indent = "  " * (level - 1)
        toc.append(add_two_spaces(f"{indent}- [{title}](#{anchor})"))

    output_path = os.path.join(output_dir, "output.md")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(toc + ["\n"] + markdown_lines))

    return output_path


# Пример использования
docx_path = "Python теория.docx"
output_dir = "MarkdownOutput"
result_path = docx_to_markdown(docx_path, output_dir)
print(f"Markdown сохранен в: {result_path}")
