# PPTemp

PPTemp is a wrapper for 
[python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html).
It enables you to make powerpoint files in a simple python commands.


## Installation

```bash
pip install pptemp
```

## Usage
```python
from pptemp import pptemp

# Initialization
presentation = pptemp()

# Initialization with template
presentation = pptemp(template="./sample/template.pptx")
    
# Title
presentation.add_title_slide("Title", "Subtitle")

# Create Blank Slide with title on the top
presentation.add_content_slide("Title of the slide")

# Create slides from figures
presentation.add_figure_slide()

# Create slides from figures with label
# Set use_bar=False if you don't want the bars to appear
presentation.add_figure_label_slide(dir_path="./sample/fig/*/")

presentation.add_figure_label_slide(dir_path="./sample/fig/*/", use_label=False)

# Save
presentation.save("./test.pptx")
```

## add_figure_label_slide()
add_figure_slide() and add_figure_label_slide() are use to import figures automatically from the "./fig" directory.

It will search figures specified by dir_path and img_path.
```
dir_path = "./fig/*/"
img_path = "*.png"
```

By Default, the title of the slides are taken from the dir_path and img_path."_" and "." are used as a separator.

```
If dir_path = "./fig/01_test/", then the title will be "test".
If img_path = "01_test.png", then the label will be "test".
```

To change where to look for the title and label, you can use the following arguments.

```
file_regex = re.compile(r".*[_/\\](.*)\.[a-zA-Z]+")
dir_regex = re.compile(r".*[_/\\](.*)[/\\]")
```

## Samples

```python
# Basic sample
python sample1.py

# Samples using template-slides
# You need to prepare template pptx files without any slides. If there is a slide, new slides will be appended.
python sample2_using_template.py
```