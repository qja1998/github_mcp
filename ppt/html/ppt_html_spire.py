from spire.presentation.common import *
from spire.presentation import *

CUR_PATH = os.path.dirname(os.path.realpath(__file__))

inputFile ="C:/Users/kwon/Downloads/sample2.pptx"
outputFile = os.path.join(CUR_PATH, "spire_result/ToHTML{i}.html")

# # Create a Presentation instance
# ppt = Presentation()

# # Load a PowerPoint document
# ppt.LoadFromFile(inputFile)

# #Save the document to HTML format
# ppt.SaveToFile(outputFile, FileFormat.Html)
# ppt.Dispose()


# Create a Presentation instance
ppt = Presentation()

# Load a PowerPoint document
ppt.LoadFromFile(inputFile)

# Get the second slide

for i, slide in enumerate(ppt.Slides):
    # Save the slide to HTML format
    slide.SaveToFile(outputFile.format(i=i), FileFormat.Html)
    ppt.Dispose()