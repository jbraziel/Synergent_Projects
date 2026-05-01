from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import re


TEMPLATE_PATH = "ACH_Auto_Proposal_Template.docx"
OUTPUT_PATH = "Generated_ACH_Auto_Proposal.docx"

FONT_NAME = "Montserrat"

TARGET_BULLET_FONT_SIZE = Pt(9)
COMPONENT_BULLET_FONT_SIZE = Pt(10)
SMALL_TEXT_SIZE = Pt(8)


proposal_data = {
    "{{proposal_name}}": "ACH Auto Loan Recapture Campaign Proposal",
    "{{proposal_date}}": "April 25, 2026",
    "{{creditunion_name}}": "Sample Credit Union",

    "{{target_conversion_rate}}": "2%",
    "{{total_targets}}": "2,047",

    "{{conversions}}": "41",
    "{{avg_auto_balance}}": "$18,500",
    "{{amount_refinanced}}": "$758,500",
    "{{first_year_interest}}": "$45,510",
    "{{auto_interest_rate}}": "6.00%",
    "{{auto_term_years}}": "5",

    "{{total_targets_2}}": "4,094",
    "{{total_targets_4}}": "8,188",

    "{{campaign_1_cost}}": "$4,950",
    "{{campaign_2_cost}}": "$8,900",
    "{{campaign_2_per_cost}}": "$4,450",
    "{{campaign_4_cost}}": "$16,020",
    "{{campaign_4_per_cost}}": "$4,005",
}


target_segments = [
    (910, "members making an ACH auto loan payment to another financial institution"),
    (472, "members making a recurring ACH payment between $400–$800 to another financial institution"),
    (628, "members who paid off their auto loan in the last 12 months"),
    (245, "members due to pay off their auto loan in the next 12 months"),
    (2047, "checking members who have a loan but no auto loan with the credit union"),
]


campaign_components = [
    "Creative concept, strategy and design",
    "Preliminary data analysis",
    "Custom programmed data extract for mailing",
    "Proofing and testing",
    "Tracking, monitoring and reporting",
    "Mailing preparation and presorting",
    "Content development",
    "Responsive email template development",
    "Digital graphic assets / social media graphics",
    "5.5” x 8.5” full color variable postcards",
    "Unique URL and QR Code redirect for 12 months",
    "Optional call file for personal outreach / follow-up",
]


def set_run_font(run, size=None, bold=None):
    run.font.name = FONT_NAME
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    if size is not None:
        run.font.size = size

    if bold is not None:
        run.bold = bold


def replace_text_in_paragraph(paragraph, replacements):
    full_text = "".join(run.text for run in paragraph.runs)

    if "{{" not in full_text:
        return

    replaced = False

    for placeholder, value in replacements.items():
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, value)
            replaced = True

    if replaced:
        paragraph.clear()
        run = paragraph.add_run(full_text)
        set_run_font(run)


def replace_simple_placeholders_in_container(container, replacements):
    for paragraph in container.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)

    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_simple_placeholders_in_container(cell, replacements)


def replace_simple_placeholders(doc, replacements):
    replace_simple_placeholders_in_container(doc, replacements)

    for section in doc.sections:
        replace_simple_placeholders_in_container(section.header, replacements)
        replace_simple_placeholders_in_container(section.footer, replacements)


def insert_bullets_at_placeholder(
    doc,
    placeholder,
    bullet_items,
    bold_first_value=False,
    font_size=TARGET_BULLET_FONT_SIZE,
    spacing_after=Pt(1),
):
    def process_container(container):
        for paragraph in container.paragraphs:
            if placeholder in paragraph.text:
                parent = paragraph._element.getparent()
                index = parent.index(paragraph._element)

                parent.remove(paragraph._element)

                for item in reversed(bullet_items):
                    new_paragraph = Document().add_paragraph()
                    new_paragraph.paragraph_format.space_after = spacing_after
                    new_paragraph.paragraph_format.space_before = Pt(0)

                    if bold_first_value:
                        quantity, sentence = item

                        bullet_run = new_paragraph.add_run("• ")
                        set_run_font(bullet_run, font_size, bold=False)

                        qty_run = new_paragraph.add_run(f"{quantity:,} ")
                        set_run_font(qty_run, font_size, bold=True)

                        sentence_run = new_paragraph.add_run(sentence)
                        set_run_font(sentence_run, font_size, bold=False)

                    else:
                        run = new_paragraph.add_run(f"• {item}")
                        set_run_font(run, font_size, bold=False)

                    parent.insert(index, new_paragraph._element)

                return True

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    if process_container(cell):
                        return True

        return False

    return process_container(doc)


def insert_total_targets_line(doc, placeholder, total_targets_text):
    def process_container(container):
        for paragraph in container.paragraphs:
            if placeholder in paragraph.text:
                paragraph.clear()

                label_run = paragraph.add_run("Total Targets (de-duped by SSN): ")
                set_run_font(label_run, size=SMALL_TEXT_SIZE, bold=False)

                value_run = paragraph.add_run(total_targets_text)
                set_run_font(value_run, size=SMALL_TEXT_SIZE, bold=True)

                return True

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    if process_container(cell):
                        return True

        return False

    return process_container(doc)


def format_conversion_objective_line(doc, placeholder, value):
    def process_container(container):
        for paragraph in container.paragraphs:
            if placeholder in paragraph.text and "Deliver Measurable Campaign Performance" in paragraph.text:
                paragraph.clear()

                run1 = paragraph.add_run("Deliver Measurable Campaign Performance ")
                set_run_font(run1, size=Pt(9), bold=True)

                run2 = paragraph.add_run("by attaining a ")
                set_run_font(run2, size=Pt(9), bold=False)

                run3 = paragraph.add_run(value)
                set_run_font(run3, size=Pt(9), bold=True)

                run4 = paragraph.add_run(
                    " conversion rate within the targeted segment, translating into increased loan originations, "
                    "recaptured balances, and improved return on marketing investment."
                )
                set_run_font(run4, size=Pt(9), bold=False)

                return True

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    if process_container(cell):
                        return True

        return False

    return process_container(doc)


def format_acceptance_line(doc, placeholder, value):
    def process_container(container):
        for paragraph in container.paragraphs:
            if placeholder in paragraph.text and "If accepted, this proposal will serve as an addendum" in paragraph.text:
                paragraph.clear()

                full_text = (
                    "If accepted, this proposal will serve as an addendum to the "
                    + value +
                    " Marketing Services Master Service Agreement and will be governed by the terms and conditions outlined therein."
                )

                run = paragraph.add_run(full_text)
                set_run_font(run, size=Pt(10), bold=False)

                return True

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    if process_container(cell):
                        return True

        return False

    return process_container(doc)


def format_conversion_metrics_values(doc):
    metric_placeholders = [
        "{{total_targets}}",
        "{{target_conversion_rate}}",
        "{{conversions}}",
        "{{avg_auto_balance}}",
        "{{amount_refinanced}}",
        "{{first_year_interest}}",
    ]

    def process_container(container):
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        text = paragraph.text.strip()

                        if text in metric_placeholders:
                            paragraph.clear()
                            run = paragraph.add_run(proposal_data[text])
                            set_run_font(run, size=Pt(8), bold=True)

                    process_container(cell)

    process_container(doc)


def format_estimated_cost_lines(doc):
    def process_container(container):
        for paragraph in container.paragraphs:
            text = paragraph.text

            if "{{campaign_1_cost}}" in text:
                paragraph.clear()
                run1 = paragraph.add_run(
                    f"Estimate based on {proposal_data['{{total_targets}}']} Postcards & Emails = "
                )
                set_run_font(run1, size=Pt(8), bold=False)

                run2 = paragraph.add_run(proposal_data["{{campaign_1_cost}}"] + "*")
                set_run_font(run2, size=Pt(8), bold=True)

            elif "{{campaign_2_cost}}" in text:
                paragraph.clear()
                run1 = paragraph.add_run(
                    f"Estimate based on {proposal_data['{{total_targets_2}}']} Postcards & Emails = "
                )
                set_run_font(run1, size=Pt(8), bold=False)

                run2 = paragraph.add_run(proposal_data["{{campaign_2_cost}}"] + "*")
                set_run_font(run2, size=Pt(8), bold=True)

            elif "{{campaign_2_per_cost}}" in text:
                paragraph.clear()
                run1 = paragraph.add_run(proposal_data["{{campaign_2_per_cost}}"] + "*")
                set_run_font(run1, size=Pt(8), bold=True)

                run2 = paragraph.add_run(" per campaign")
                set_run_font(run2, size=Pt(8), bold=False)

            elif "{{campaign_4_cost}}" in text:
                paragraph.clear()
                run1 = paragraph.add_run(
                    f"Estimate based on {proposal_data['{{total_targets_4}}']} Postcards & Emails = "
                )
                set_run_font(run1, size=Pt(8), bold=False)

                run2 = paragraph.add_run(proposal_data["{{campaign_4_cost}}"] + "*")
                set_run_font(run2, size=Pt(8), bold=True)

            elif "{{campaign_4_per_cost}}" in text:
                paragraph.clear()
                run1 = paragraph.add_run(proposal_data["{{campaign_4_per_cost}}"] + "*")
                set_run_font(run1, size=Pt(8), bold=True)

                run2 = paragraph.add_run(" per campaign")
                set_run_font(run2, size=Pt(8), bold=False)

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_container(cell)

    process_container(doc)


def make_paragraph_small(paragraph):
    for run in paragraph.runs:
        set_run_font(run, size=SMALL_TEXT_SIZE)


def format_small_text_sections(doc):
    target_phrases = [
        "Total Targets (de-duped by SSN):",
        "(Auto loan figures are based on",
    ]

    def process_container(container):
        for paragraph in container.paragraphs:
            text = paragraph.text.strip()
            for phrase in target_phrases:
                if text.startswith(phrase):
                    make_paragraph_small(paragraph)
                    break

        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_container(cell)

    process_container(doc)


def build_split_placeholder_pattern(placeholder):
    return r"\s*".join(re.escape(char) for char in placeholder)


def replace_placeholders_in_all_xml_text(doc, replacements):
    def get_text_nodes():
        return list(doc.element.xpath(".//w:t"))

    def replace_one_placeholder(placeholder, value):
        pattern = build_split_placeholder_pattern(placeholder)

        while True:
            text_nodes = get_text_nodes()
            full_text = ""
            char_map = []

            for node_index, node in enumerate(text_nodes):
                node_text = node.text or ""

                for char_index, char in enumerate(node_text):
                    full_text += char
                    char_map.append((node_index, char_index))

            match = re.search(pattern, full_text)

            if not match:
                break

            start_pos = match.start()
            end_pos = match.end() - 1

            start_node_index, start_char_index = char_map[start_pos]
            end_node_index, end_char_index = char_map[end_pos]

            start_node = text_nodes[start_node_index]
            end_node = text_nodes[end_node_index]

            if start_node_index == end_node_index:
                original = start_node.text or ""
                start_node.text = (
                    original[:start_char_index]
                    + str(value)
                    + original[end_char_index + 1:]
                )
            else:
                start_text = start_node.text or ""
                end_text = end_node.text or ""

                prefix = start_text[:start_char_index]
                suffix = end_text[end_char_index + 1:]

                start_node.text = prefix + str(value) + suffix

                for i in range(start_node_index + 1, end_node_index + 1):
                    text_nodes[i].text = ""

    for placeholder, value in replacements.items():
        replace_one_placeholder(placeholder, value)


def find_unreplaced_placeholders(doc):
    full_text = []

    for text_node in doc.element.xpath(".//w:t"):
        if text_node.text:
            full_text.append(text_node.text)

    joined_text = "\n".join(full_text)
    return sorted(set(re.findall(r"\{\{[\s\S]*?\}\}", joined_text)))


def main(output_path=None):
    doc = Document(TEMPLATE_PATH)

    proposal_data["{{target_segment_summary}}"] = "\n".join(
        f"• {qty:,} {sentence}" for qty, sentence in target_segments
    )

    proposal_data["{{campaign_components}}"] = "\n".join(
        f"• {item}" for item in campaign_components
    )

    format_conversion_objective_line(
        doc,
        "{{target_conversion_rate}}",
        proposal_data["{{target_conversion_rate}}"]
    )

    format_acceptance_line(
        doc,
        "{{creditunion_name}}",
        proposal_data["{{creditunion_name}}"]
    )

    insert_bullets_at_placeholder(
        doc,
        "{{target_segment_summary}}",
        target_segments,
        bold_first_value=True,
        font_size=TARGET_BULLET_FONT_SIZE,
        spacing_after=Pt(1),
    )

    insert_total_targets_line(
        doc,
        "{{total_targets_line}}",
        proposal_data["{{total_targets}}"]
    )

    insert_bullets_at_placeholder(
        doc,
        "{{campaign_components}}",
        campaign_components,
        bold_first_value=False,
        font_size=COMPONENT_BULLET_FONT_SIZE,
        spacing_after=Pt(4),
    )

    format_conversion_metrics_values(doc)
    format_estimated_cost_lines(doc)

    replace_simple_placeholders(doc, proposal_data)
    replace_placeholders_in_all_xml_text(doc, proposal_data)
    format_small_text_sections(doc)

    remaining = find_unreplaced_placeholders(doc)

    if output_path is None:
        output_path = OUTPUT_PATH

    doc.save(output_path)

    print(f"Proposal created successfully: {output_path}")

    if remaining:
        print("\nThese placeholders are still unreplaced:")
        for item in remaining:
            print(f" - {item}")
    else:
        print("All placeholders were replaced.")


if __name__ == "__main__":
    main()
