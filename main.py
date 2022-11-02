from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from os import listdir, makedirs
from os.path import isfile, join, exists, dirname, splitext
from datetime import datetime as d
import constants
import sys


# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    application_path = dirname(sys.executable)
elif __file__:
    application_path = dirname(__file__)

output_dir = join(application_path, constants.OUTPUT_DIR)
image_dir = join(application_path, constants.IMAGE_DIR)
template_dir = join(application_path, constants.TEMPLATE_DIR)
template_file_path = join(template_dir, constants.TEMPLATE_FILE_NAME)

date = d.now()
current_datetime = date.strftime("%Y-%m-%d %H:%M:%S")


is_exist = exists(output_dir)

if not is_exist:
    makedirs(output_dir)


image_files = [f for f in listdir(image_dir) if isfile(join(image_dir, f)) and  not f.startswith('.')]

sorted_image_files = sorted(image_files)

paginated_sorted_image_files = [sorted_image_files[i:i+4]
                                for i in range(0, len(sorted_image_files), 4)]


prs = Presentation(template_file_path)

# Adding intro slides

note_slide = prs.slides
for slide in note_slide:
    slide.shapes.title.text = "generated image slides"

    slide.placeholders[1].text = current_datetime


custom_slide_layout = prs.slide_layouts[11]

for item_array in paginated_sorted_image_files:
    slide = prs.slides.add_slide(custom_slide_layout)
    try:
        images_placeholder_idx_arr = []
        text_placeholder_idx_arr = []

        for shape in slide.placeholders:
            if shape.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                images_placeholder_idx_arr.append(
                    {"idx": shape.placeholder_format.idx, "name": shape.name})
            if shape.placeholder_format.type == PP_PLACEHOLDER.OBJECT:
                text_placeholder_idx_arr.append(
                    {"idx": shape.placeholder_format.idx, "name": shape.name})
            # print('%d %s %s' % (shape.placeholder_format.idx,
            #       shape.name, shape.placeholder_format.type))

        images_placeholder_idx_arr_sorted_by_name = sorted(
            images_placeholder_idx_arr, key=lambda x: x['name'])
        text_placeholder_idx_arr_sorted_by_name = sorted(
            text_placeholder_idx_arr, key=lambda x: x['name'])

        for index in range(0, len(item_array)):
            # print(idx)
            image_name = item_array[index]
            image_path = join(image_dir, image_name)

            placeholder = slide.placeholders[images_placeholder_idx_arr_sorted_by_name[index]['idx']]
            picture = placeholder.insert_picture(image_path)

            placeholder_text = slide.placeholders[text_placeholder_idx_arr_sorted_by_name[index]['idx']]
            placeholder_text.text = splitext(image_name)[0]

    except Exception as e:
        print("image insertion issues")

file_name = f'{date.strftime("%Y%m%d%H%M%S")}_{constants.FILE_NAME}'
output_path = join(output_dir, file_name)

prs.save(output_path)
