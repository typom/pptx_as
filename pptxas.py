'''

'''
import win32com.client as wc


class Presentation:

    app = None

    def __init__(self):
        app = Presentation.app
        if app is None:
            app = wc.Dispatch("PowerPoint.Application")
        app.Visible = True
        self.presentation_ = app.Presentations.Add()
        self.slides = []

    def add_slide(self, title=None, position=None, **kwargs):
        if position is None:
            position = len(self.slides)+1
        position = min(len(self.slides)+1, position)
        new_slide = Slide(self, title, position, **kwargs)
        self.slides.append(new_slide)
        return new_slide


class Slide:
    def __init__(self, presentation, title=None, position=1, **kwargs):
        self.slide_ = presentation.presentation_.Slides.Add(position, 12)
        if title is not None:
            self.draw_textbox(title, 10, 10)

    def draw_textbox(self, text, x, y):
        shape = self.slide_.Shapes.AddShape(1, x, y, 300, 50)
        # Add text: bank name (asset size in millions)
        shape.TextFrame.TextRange.Text = text
        # Reduce left and right margins
        # shape.TextFrame.MarginLeft = shape.TextFrame.MarginRight = 0
        # Use 12pt font
        shape.TextFrame.TextRange.Font.Size = 12

    def getSize(self):
        return self.slide_.get_size()


p = Presentation()
s1 = p.add_slide(title='slide 1')
s2 = p.add_slide(title='slide 2')
