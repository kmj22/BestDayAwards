import io

import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

ROOT = 'files/2024'
AWARD_NAME_FONT = {
    'name': 'ScholarlyAmbition-Regular',
    'file': 'ScholarlyAmbition-Regular.ttf'
}
TEAM_NAME_FONT = {
    'name': 'Trattatello',
    'file': 'Trattatello.ttf'
}

def add_font(fnt):
    pdfmetrics.registerFont(TTFont(fnt['name'], fnt['file']))


def fill_pdf_template(template_path, output_path, teams, stopAt=99):
    add_font(TEAM_NAME_FONT)
    add_font(AWARD_NAME_FONT)

    output_pdf = PdfWriter()

    total = len(teams)
    i = 0

    for team in teams:
        i += 1

        print(f'Working on {i}/{total}')
        new_pdf = PdfReader(template_path)
        page = new_pdf.pages[0]

        teamName = team[0]
        awardName = team[1]
        squeeze = False
        if len(team) > 2:
            squeeze = True

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(11*inch, 8.5*inch))

        # Grid goes >^
        # award
        awardFont = AWARD_NAME_FONT['name']
        awardSize = 44

        leftX = 3.22*inch
        lineWidth = 7.26*inch
        textWidth = stringWidth(awardName, awardFont, awardSize)

        message = None
        # Repeat until it fits
        verticalSpace = 2*inch
        if squeeze:
            verticalSpace = awardSize
        while True:
            can.setFont(awardFont, awardSize)
            message_style = ParagraphStyle('Normal',
                fontName=awardFont,
                fontSize=awardSize,
                alignment=TA_CENTER,
                leading=awardSize
            )
            message = Paragraph(awardName, style=message_style)
            w, h = message.wrap(lineWidth, verticalSpace)

            if h < verticalSpace:
                break
            else:
                awardSize -= 1

        x = leftX
        y = 2.9*inch
        message.drawOn(can, x, y)

        # team
        teamFont = TEAM_NAME_FONT['name']
        teamSize = 28

        x = 6.5*inch
        y = 5.2*inch
        lineWidth = 4.1*inch

        while stringWidth(teamName, teamFont, teamSize) > lineWidth:
            teamSize -= 1
            y += 1

        can.setFont(teamFont, teamSize)
        can.drawString(x, y, teamName)

        # DONE- save it!
        can.save()

        # move to the beginning of the StringIO buffer
        packet.seek(0)

        # create a new PDF with Reportlab
        name_pdf = PdfReader(packet)

        print('merge_page()')
        page.merge_page(name_pdf.pages[0])
        print('compress...')

        # This line helps prevent "PyPDF2.utils.PdfReadError: Xref table not zero-indexed" error
        page.compress_content_streams()

        # Add the filled page to the output PDF
        print('Adding page...')
        output_pdf.add_page(page)

        if i > stopAt:
            break

    # Save the filled PDF to the output path
    with open(output_path, 'wb') as output_file:
        print('saving...')
        output_pdf.write(output_file)


def get_test_teams():
    return [
        ['Team with a really really really really long name', 'Most independent team'],
        ['Team with a really really really really long name', 'Most independent team test test test test test test'],
        ['Team with a really really really really long name', '#1 Highest Energy Output to Cost Ratio, Overall Excellence Award'],
        ['Team Name', '#1 Highest Energy Output to Cost Ratio,<br />Overall Excellence Award'],
    ]


def get_name_award_combos():
    # Read CSV data
    data = pd.read_csv(csv_path, names=['award', 'team', 'squeeze'])
    teams = []
    for index, row in data.iterrows():
        data = [row[0], row[1]]
        if (len(row) > 2 and row[2] == 1):
            data.append(1)
        teams.append(data)
    return teams


# Need: template PDF, csv with group names + award names
# You can add "<br />" to the award names to manually add line breaks if the auto ones don't look good
# Make sure the group/award names don't have any commas!  They'll mess up the CSV

# If fonts need to change:
# Install the font from the internet (probably a TTF file)
# By default the pdfmetrics.registerFont() method will check logical default system folders, like:
# System/Library/Fonts (+ /Supplemental)
if __name__ == "__main__":
    # Specify the paths to your template PDF and CSV file
    template_path = f'{ROOT}/template.pdf'
    csv_path = f'{ROOT}/data.csv'
    output_path = f'{ROOT}/awards.pdf'

    teams = get_name_award_combos()

    # Fill the PDF template with data and save the output PDF
    fill_pdf_template(template_path, output_path, teams)
