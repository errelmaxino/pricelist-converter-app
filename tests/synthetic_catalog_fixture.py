"""Invented catalog samples with independently written expected normalized records.

These cases are NOT OEM data and do NOT establish vehicle-part compatibility.
The sample PDF is generated in memory; no real supplier file is needed.
"""
from __future__ import annotations

CASES = [
    ("VK-101", "TOY COROLLA '07", [("TOYOTA", "COROLLA", "2007")]),
    ("VK-102", "MIT LANCER '08", [("MITSUBISHI", "LANCER", "2008")]),
    ("VK-103", "MAZ BONGO 2000", [("MAZDA", "BONGO 2000", "")]),
    ("VK-104", "GENERIC 2007", [("", "GENERIC", "2007")]),
    ("VK-105", "TOY COROLLA '07 / MIT LANCER '08",
     [("TOYOTA", "COROLLA", "2007"), ("MITSUBISHI", "LANCER", "2008")]),
    ("VK-106", "NIS SENTRA 1998", [("NISSAN", "SENTRA", "1998")]),
    ("VK-107", "TOY COROLLA '07", [("TOYOTA", "COROLLA", "2007")]),
    ("VK-107", "MIT LANCER '08", [("MITSUBISHI", "LANCER", "2008")]),
]


def pdf_with_lines(lines: list[str]) -> bytes:
    """Make a deliberately simple, text-searchable, one-page PDF using stdlib."""
    from io import BytesIO
    commands = ["BT /F1 10 Tf 40 750 Td"]
    for i, line in enumerate(lines):
        if i:
            commands.append("0 -22 Td")
        escaped = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
        commands.append(f"({escaped}) Tj")
    commands.append("ET")
    stream = "\n".join(commands).encode("latin-1")
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
    ]
    b = bytearray(b"%PDF-1.4\n")
    positions = []
    for i, obj in enumerate(objs, 1):
        positions.append(len(b))
        b.extend(f"{i} 0 obj\n".encode() + obj + b"\nendobj\n")
    root = len(b)
    b.extend(b"xref\n0 6\n0000000000 65535 f \n")
    for pos in positions:
        b.extend(f"{pos:010} 00000 n \n".encode())
    b.extend(f"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{root}\n%%EOF\n".encode())
    return bytes(b)


def generated_pdf() -> bytes:
    return pdf_with_lines([f"{code} Front {description} {100+i*10:.2f}"
                           for i, (code, description, _expected) in enumerate(CASES)])


def expected_records() -> list[tuple[str, str, str, str]]:
    return [(code, brand, model, year)
            for code, _desc, records in CASES
            for brand, model, year in records]
