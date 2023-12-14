from pptx.dml.color import RGBColor

def change_color(cell):
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(255, 0, 0)
    # paragraph =  cell.text_frame.paragraphs[0]
    # paragraph.font.color.rgb = RGBColor(255, 0, 0)
    # paragraph.font.bold = True