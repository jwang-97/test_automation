from pptx import Presentation
import pptx, pptx.util, glob, os
# import scipy.misc
# from pptx.util import Inches
# import imageio
import cv2
import configparser


def contact_pic_hooks():
    cf = configparser.ConfigParser()
    cf.read(os.getcwd() + r'\auto_test_config.ini', encoding='utf-8')
    OUTPUT_TAG = cf.get("merge_template", "dirct")
    prs = Presentation(OUTPUT_TAG)
    # open
    # prs_exists = Presentation(r"C:\Users\JWang294\src\cis\libcalc\auto_report_pipeline1.1\60095160_RevA_CPC_Template_v4.04.pptx")

    # default slide width
    #prs.slide_width = 9144000
    # slide height @ 4:3
    #prs.slide_height = 6858000
    # slide height @ 16:9
    # prs.slide_height = 5143500
    # title slide
    # slide = prs.slides.add_slide(prs.slide_layouts[0])
    # blank slide
    #slide = prs.slides.add_slide(prs.slide_layouts[6])
    # set title
    # slide = prs.slides.add_slide(prs.slide_layouts[0])
    # title = slide.shapes.title
    # subtitle = OUTPUT_TAG
    
    pic_left = int(prs.slide_width*0.15)
    pic_top = int(prs.slide_height*0.1)
    pic_width = int(prs.slide_width*0.7)
    pic_height = int(prs.slide_height*0.5)
    for g in glob.glob(r'C:\Users\JWang294\Documents\projectfile\test_config_cases\test-20210416173550-lbo-0416cpc\Contact\contact_script_test\utility_files\*.jpg'):
        print (g)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        # slide.shapes.title.textframe.textrange.text = "test contact slide"
        slide.shapes.title.text = "test contact slide"
        tb = slide.shapes.add_textbox(0.53, 2.39, 10.36, 1.32)
        p =  tb.text_frame.add_paragraph()
        p.text = "ROP Controlled: " + "50" + " ft/hr"
        p.font.size = pptx.util.Pt(25)

        # img = imageio.imread(g, pilmode="RGB")
        # img = imageio.imread(g, plugin='matplotlib')
        img = cv2.imread(g)
        # pic_height = int(pic_width * img.shape[0] / img.shape[1])
        #pic   = slide.shapes.add_picture(g, pic_left, pic_top)
        # pic = slide.shapes.add_picture(g, pic_left, pic_top, pic_width, pic_height)
        pic1 = slide.shapes.add_picture(g, 2.4, 5.87, 16.59, 14.13)
        pic2 = slide.shapes.add_picture(g, 22.19, 5.87, 16.59, 14.13)

    # prs.save("%s.pptx" % OUTPUT_TAG)
    prs.save(OUTPUT_TAG)
if __name__ == "__main__":
    contact_pic_hooks()