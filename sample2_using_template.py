import pptemp
from datetime import date

if __name__ == '__main__':
    
    # initialization
    presentation = pptemp.pptemp("./sample/sample_template.pptx")
        
    # Slide 1 Title
    slide = presentation.add_title_slide("Importing Figure", str(date.today()))
           
    # Create slides from figures with label
    # Set use_bar=False if you don't want the bars to appear
    presentation.add_figure_label_slide(dir_path="./sample/fig/*/", use_bar=False)
    
    # Create slides from figures without label
    # Set use_bar=False if you don't want the bars to appear
    presentation.add_figure_slide(dir_path="./sample/fig/*/", use_bar=False)
        
    # save
    presentation.save("./sample_output/sample2_using_template.pptx")
    