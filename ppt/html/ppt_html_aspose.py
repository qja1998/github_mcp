import os
import aspose.slides as slides

CUR_PATH = os.path.dirname(os.path.realpath(__file__))
print(CUR_PATH)

# Load the presentation file
pres = slides.Presentation("C:/Users/kwon/Downloads/sample2.pptx")

# Create HTML options
options = slides.export.HtmlOptions()

# Create a responsive HTML controller
controller = slides.export.ResponsiveHtmlController() 

# Set controller as HTML formatter
options.html_formatter = slides.export.HtmlFormatter.create_custom_formatter(controller)

# Save as HTML
pres.save(os.path.join(CUR_PATH, "aspose_result/ToHTML.html"), slides.export.SaveFormat.HTML, options)