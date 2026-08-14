import unicodedata

from .ranges import column_name


def text_width(value):
    total = 0
    for char in str(value):
        total += 2 if unicodedata.east_asian_width(char) in ("W", "F") else 1
    return total


def pad(value, width):
    value = str(value)
    return value + " " * max(0, width - text_width(value))


def render_table(rows, row_start, col_start):
    labels = [column_name(col_start + index + 1) for index in range(len(rows[0]) if rows else 0)]
    widths = [text_width(label) for label in labels]
    for row in rows:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], text_width(value))
    row_width = max(1, len(str(row_start + len(rows))))

    def border():
        return "+" + "-" * (row_width + 2) + "+" + "".join("-" * (width + 2) + "+" for width in widths)

    result = [border()]
    result.append("| %s |%s" % (pad("#", row_width), "".join(" %s |" % pad(label, widths[i]) for i, label in enumerate(labels))))
    result.append(border())
    for offset, row in enumerate(rows):
        result.append("| %s |%s" % (
            pad(row_start + offset + 1, row_width),
            "".join(" %s |" % pad(value, widths[i]) for i, value in enumerate(row)),
        ))
        result.append(border())
    return "\n".join(result)
