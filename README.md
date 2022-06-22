# PPTemp

PPTemp is a wrapper for 
[python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html).
It enables you to make powerpoint files in a simple python commands.

## Installing Dependencies

```
pip install -r requirement.txt
```

## Usage
```python
import pptemp

# Initialization
presentation = pptemp.pptemp()
    
# Title
presentation.add_title_slide("Title", "Subtitle")

# Create Blank Slide with title on the top
presentation.add_title_slide("Title of the slide")

# Create slides from figures with label
presentation.add_figure_label_slide()

# Create slides from figures
presentation.add_figure_slide()

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