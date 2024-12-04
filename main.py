import os
from functions import *

def main_function(slides_id):
    # Set up logging
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Load environment variables
    load_dotenv()
    
    # Extract slides_id from the event
    # slides_id = event.get('slides_id')
    if not slides_id:
        logger.error('slides_id not provided')
        return {'statusCode': 400, 'body': 'slides_id not provided'}
    
    try:
        # Get slide data
        slide_data_respone = get_slide_data(slides_id)

        prompt_added = slide_data_respone[0].get("prompt")
        tone_requested = slide_data_respone[0].get("tone")
        template_type = slide_data_respone[0].get("template_type")
        slides_count = slide_data_respone[0].get("slides_count")
        slides_count = slides_count_to_int(slides_count)

        # Gather the text 
        downloaded_file_path = download_file(slide_data_respone)

        text_content = slide_data_respone[0].get("text_content")

        total_text_uploaded = extract_text(downloaded_file_path) + extract_text(text_content)

        # Download the images
        image_directory = "/tmp/images"
        if downloaded_file_path is not None:
            extract_images(downloaded_file_path, image_directory, max_images=10)

        template_type = slide_data_respone[0].get("template_type")

        template_category_meta_data = get_template_category_meta_data(template_type)

        picked_presentation_template = presentation_picker_from_meta_data(total_text_uploaded, prompt_added, tone_requested, template_category_meta_data, template_type)

        picked_presentation_template = "/tmp/" + "template.pptx"

        # get to know template
        template_information_parsed = parse_template_slides(picked_presentation_template)

        presentation_plan_structure = generate_presentation_plan(total_text_uploaded, template_information_parsed, slides_count)

        structured_placeholders_to_be_filled = get_slide_placeholders(presentation_plan_structure, template_information_parsed)

        presentation_content = generate_presentation_content(total_text_uploaded, prompt_added, tone_requested, structured_placeholders_to_be_filled)

        presentation_content = add_image_info_to_presentation_content(presentation_content, template_information_parsed)
        print(presentation_content)

        image_folder = "/tmp/images"
        image_assignment = assign_images(image_folder, presentation_content)

        image_assignment = fetch_and_download_images(image_assignment, image_folder)

        empty_qlina__url = f"https://cdn.qlina.ai/empty_qlina.pptx"
        response = requests.get(empty_qlina__url)
        response.raise_for_status()
        with open("/tmp/empty_qlina.pptx", "wb") as file:
            file.write(response.content)

        empty_template = '/tmp/empty_qlina.pptx'

        spire_template = generate_presentation_from_plan(presentation_plan_structure, picked_presentation_template, empty_template)

        output_file = '/tmp/spire_template_clean.pptx'

        spire_template_clean = remove_evaluation_warnings(spire_template, output_file)

        save_as_almost_final = '/tmp/final_presentation_qlina_not_yet_images.pptx'
        insert_content_into_presentation(spire_template_clean, presentation_content, save_as_almost_final)

        save_as_final = '/tmp/final_presentation_qlina.pptx'
        replace_images_in_presentation(save_as_almost_final, image_assignment, save_as_final, image_folder)

        print(prompt_added)

        # push pdf and pptx
        upload_final_to_bunny(save_as_final, slides_id)

        initaite_pptx_2_pdf(slides_id)

        # Extract the title from the first slide's content
        title = presentation_content['slides'][0].get("contents", [""])[0]

        # update status 
        response = update_supabase(slides_id, title)
        print(response)

        # clean up tmp folder
        shutil.rmtree("/tmp/")
        
        return {'statusCode': 200, 'body': 'QLINA Processing completed successfully'}
    
    except Exception as e:
        logger.exception('Error processing slides_id %s', slides_id)
        return {'statusCode': 500, 'body': str(e)}

main_function(os.environ.get("SLIDES_ID", ""))
