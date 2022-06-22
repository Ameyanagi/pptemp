import pptemp
from datetime import date

if __name__ == '__main__':
    
    # initialization
    # presentation = pptemp.pptemp("./template.pptx")
    presentation = pptemp.pptemp()
        
    # Slide 1 Title
    slide = presentation.add_title_slide("Importing Figure", str(date.today()))
           
    # Create slides from figures with label
    presentation.add_figure_label_slide()
    
    # Create slides from figures with label
    presentation.add_figure_slide()
    
    # save
    presentation.save("./test.pptx")
    