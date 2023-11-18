import io
import time
import asyncio
from docxtpl import DocxTemplate
import base64
import concurrent.futures
import streamlit as st
from openai import OpenAI

# # Define your variables
# data = {
#     'deployment_constraints': 'abc',
#     'energy_analysis_field': 'def',
#     'citation_count': "3",
#     'expert_name': 'Dr. Memon',
#     'degree_field': 'Computer Science',
#     'degree_university': 'XYZ University',
#     'degree_year': 2015,
#     'field_of_expertise': 'Computer Science',
#     'proposed_endeavor_description': 'develop cutting-edge machine learning algorithms',
#     'proposed_endeavor': 'developing state-of-the-art machine learning algorithms',
#     'employment_position': 'Research Scientist',
#     'employment_employer': 'ABC Research Lab',
#     'research_applications': 'the development of advanced machine learning applications',
#     'benefits': 'revolutionizing the field of machine learning',
#     'sources': 'industry reports and expert testimonials',
#     'num_testimonials': 2,
#     'testimonial_1': 'This expert\'s work has greatly impacted our understanding of AI.',
#     'testimonial_author_1': 'John Doe',
#     'testimonial_author_position_1': 'CEO',
#     'testimonial_author_organization_1': 'AI Tech Inc.',
#     'testimonial_2': 'Their research is a game-changer in the field of machine learning.',
#     'testimonial_author_2': 'Jane Smith',
#     'testimonial_author_position_2': 'Director of Research',
#     'testimonial_author_organization_2': 'Tech Innovations Ltd.',
#     'analyzing_organization': 'Tech Analysis Group',
#     'global_race_description': 'a global competition for technological superiority',
#     'advanced_applications': 'autonomous vehicles to personalized medicine',
#     'energy_analysis_organization': 'GreenTech Solutions',
#     'news_source': 'Tech Insights Magazine',
#     'key_energy_issues': 'lack of sustainable energy sources',
#     'transmission_limitations': 'a restricted range of data transmission',
#     'innovative_techniques': 'novel deep learning architectures',
#     'additional_sources': '[AI Journal, Energy Reports]',
#     'research_topics': 'AI ethics and interpretable machine learning',
#     'development_description': 'a novel interpretable machine learning model',
#     'notable_citations': '200',
#     'num_journal_articles': '20',
#     'num_first_authored': '15',
#     'ms_field': 'Artificial Intelligence',
#     'ms_university': 'LMN University',
#     'phd_field': 'Machine Learning',
#     'phd_university': 'LMN University',
#     'current_position': 'Associate Professor',
#     'current_position_university': 'XYZ University',
#     'current_position_location': 'Cityville, USA',
#     'proposed_systems': 'remote patient monitoring and personalized health tracking',
#     'diabetes_population': '500 million',
#     'diabetes_percentage': '10%',
#     'projected_diabetes_population': '700 million',
#     'projected_year': '2030',
#     'testimonial_1': 'This expert\'s work has greatly impacted our understanding of AI.',
#     'testimonial_author_1': 'John Doe',
#     'testimonial_author_position_1': 'CEO',
#     'degree_evaluation': "xyz"
# }

data = {
    'deployment_constraints': 'The proposed endeavor should adhere to ethical and legal standards.',
    'energy_analysis_field': 'Energy consumption analysis in machine learning models.',
    'citation_count': 45,
    'petitioner_name': 'Dr. Sarah Johnson',
    'degree_field': 'Computer Science',
    'degree_university': 'Stanford University',
    'degree_year': 2015,
    'field_of_expertise': 'Machine Learning',
    'proposed_endeavor_description': 'Developing cutting-edge machine learning algorithms for real-world applications.',
    'proposed_endeavor': 'Developing state-of-the-art machine learning algorithms',
    'employment_position': 'Research Scientist',
    'employment_employer': 'Tech Innovations Research Lab',
    'research_applications': 'The development of advanced machine learning applications.',
    'benefits': 'Advancing the field of machine learning with practical applications.',
    'sources': 'Industry reports, academic publications, and expert testimonials.',
    'num_testimonials': 1,
    'testimonials': [
        {
            'content': 'Dr. Johnson\'s work has significantly advanced our understanding of AI.',
            'author': 'Alex Rodriguez',
            'position': 'CTO',
            'organization': 'Innovate Tech Solutions'
        },
        {
            'content': 'Her research is a game-changer in the field of machine learning.',
            'author': 'Emily Davis',
            'position': 'Director of Research',
            'organization': 'Tech Dynamics Inc.'
        }
    ],
    'analyzing_organization': 'Tech Analysis Group',
    'global_race_description': 'A global competition for technological excellence in various domains.',
    'advanced_applications': 'From autonomous vehicles to personalized medicine.',
    'energy_analysis_organization': 'GreenTech Solutions',
    'news_source': 'Tech Insights Journal',
    'key_energy_issues': 'Addressing the lack of sustainable energy sources in machine learning.',
    'transmission_limitations': 'Overcoming restricted data transmission ranges.',
    'innovative_techniques': 'Implementing novel deep learning architectures.',
    'additional_sources': ['AI Journal', 'Energy Reports'],
    'notable_achievements': ['Best Paper Award', 'IEEE Conference'],
    'research_topics': 'AI ethics, interpretable machine learning, and sustainable energy solutions.',
    'development_description': 'Creating a novel interpretable machine learning model.',
    'notable_citations': 200,
    'num_journal_articles': 20,
    'num_first_authored': 15,
    'ms_field': 'Artificial Intelligence',
    'ms_university': 'MIT',
    'phd_field': 'Machine Learning',
    'phd_university': 'Stanford University',
    'current_position': 'Associate Professor',
    'current_position_university': 'Harvard University',
    'current_position_location': 'Cambridge, USA',
    'proposed_systems': 'Remote patient monitoring and personalized health tracking',
    'diabetes_population': '500 million',
    'diabetes_percentage': '10%',
    'projected_diabetes_population': '700 million',
    'projected_year': '2030',
    'degree_evaluation': "Outstanding",
    'research_area': 'Machine Learning',
    'research_subfield': 'Interpretable Machine Learning'
}

# Set your OpenAI API key
api_key = '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'

# Initialize the OpenAI client
client = OpenAI(api_key=api_key)

# Define a function to generate responses asynchronously
async def generate_response(prompt):
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a US Immigration attorney assistant having expertise in writing I-140 petition letters, presenting strong arguments for USA Employement Based 2nd preference (EB2) National Interest Waiver (NIW) petitions. Your role is to create and compose a comprehensive petition letter for a Ph.D. holder, to be written by the attorney, based on the provided information using instructions given for user content , use data variables in it"},
            {"role": "user", "content": prompt},
        ]
    )
    return response.choices[0].message.content

# Define a function to generate responses asynchronously using asyncio.gather
async def generate_responses(prompts):
    async with concurrent.futures.ThreadPoolExecutor() as executor:
        responses = await asyncio.gather(*[generate_response(prompt) for prompt in prompts])
    return responses

# Define a function to generate responses asynchronously
async def generate_document(data):
    # Your code for generating the document goes here
    # Define your prompts

#     class CoverLetterData:
#         def __init__(self, data):
#             self.introduction = (
#             f"1. **Introduction to Standard of Proof:**\n"
#             f"- Compose an introduction of up to 300 words that elaborates on the 'preponderance of the evidence' standard and how it applies to EB2 NIW petitions, drawing insights from Matter of Dhanasar, Matter of E-M-, and U.S. v. Cardoza-Fonseca"
#             f"For Example, Generate like Below"
#             f"Note that the standard of proof for petitions filed for National Interest Waiver cases is the 'preponderance of the evidence' standard. See Matter of Dhanasar, 26 I&N Dec. 884, 889 (AAO 2016). Thus, if the petitioner submits relevant, probative, and credible evidence that leads USCIS to believe that the claim is 'more likely than not' or 'probably true,' the petitioner has satisfied the standard of proof. Matter of E-M-, 20 I&N Dec. 77, 79-80 (Comm’r 1989); see also U.S. v. Cardoza-Fonseca, 480 U.S. 421 (1987) (discussing 'more likely than not' as a greater than 50% chance of an occurrence taking place).\n\n"

#             f"- Compose the Highlighting Advanced Degree of up to 500 words that Provide a compelling narrative around {data['petitioner_name']}'s advanced degree, emphasizing the significance of their Ph.D. in {data['degree_field']} from {data['degree_university']}. You have to present this information to showcase {data['petitioner_name']} as a highly qualified professional."
#             f"I. {data['petitioner_name']} is a member of the professions holding an advanced degree\n"
#             f"I. {data['petitioner_name']} received a PhD in {data['degree_field']} from {data['degree_university']} in {data['degree_year']}. As evidence of this, we are submitting copies of {data['petitioner_name']}'s diploma and transcripts (Exhibit [advanced degree, transcripts]). Since {data['petitioner_name']} completed their education outside the United States, we are also submitting a detailed advisory evaluation of their educational credentials (Exhibit {data['degree_evaluation']}). As such, {data['petitioner_name']} is qualified as a member of the professions holding an advanced degree.\n\n"

#             f"- Compose the Detailing Proposed Endeavor up to 800 words that enhance the description of {data['petitioner_name']}'s proposed endeavor of developing state-of-the-art {data['field_of_expertise']} in the field of {data['research_area']}? All aspects should be emphasized to make it stand out."
#             f"II. {data['petitioner_name']}'s proposed endeavor is in the field of {data['field_of_expertise']} to {data['proposed_endeavor_description']}\n"
#             f"As an expert in the field of {data['field_of_expertise']}, {data['petitioner_name']}'s proposed endeavor is to {data['proposed_endeavor']} (Exhibits [PES-EL]).\n"
#             f"This petition waives the job offer requirement, and the petitioner's proposed endeavor is separate from their proposed employment. However, we are submitting {data['petitioner_name']}'s plans for employment in the field to confirm their commitment and capacity to advance their proposed endeavor. Based on their education and research background, {data['petitioner_name']} plans to be employed as a {data['employment_position']} at {data['employment_employer']} or a similar employer (Exhibit [PES-EL]). {data['petitioner_name']} intends to continue their research on {data['proposed_endeavor']} (Exhibits [PES-EL]). This said, the focus of this prong should be on the proposed endeavor itself rather than {data['petitioner_name']}'s employment.\n\n"

#             f"- Compose the “Emphasizing Merit and National Importance” up to 800 words that show Given {data['petitioner_name']}'s research applications and the testimonials provided, you have to strengthen the argument for the substantial merit and national importance of his proposed endeavor. Discuss specific examples or points that should be highlighted."
#             f"{data['petitioner_name']}'s proposed endeavor of {data['proposed_endeavor']} has both substantial merit and national importance"
#             f"{data['petitioner_name']}'s research related to {data['proposed_endeavor']} has great substantial merit and national importance. Among other applications, their research is relevant to {data['research_applications']} (Exhibits 1-4, [publications, testimonial letters])."
#             f"i. {data['petitioner_name']}'s proposed endeavor has substantial merit"
#             f"{data['petitioner_name']}'s research advancing their proposed endeavor is of great importance because it allows for {data['benefits']} (Exhibits 1-4, {data['sources']}). Fellow experts in the field provide additional insight into the merit of this endeavor:"
#             f"{data['petitioner_name']}'s research advancing his proposed endeavor is of great importance because it allows for improved healthcare to global communities. It is estimated that there are currently {data['diabetes_population']} individuals afflicted with diabetes globally, or {data['diabetes_percentage']} of the world’s population. That number is projected to grow to approximately {data['projected_diabetes_population']} by {data['projected_year']}. Health officials encourage diabetics to monitor their blood sugar daily. Finger sticks, in which the individual pricks his or her fingertip to provide blood for analysis by a portable sensor, remains the primary method for this test. However, finger sticks have several limitations, including the pain of repeated jabs to fingertips, as well as difficulty of application on individuals whose hands are swollen, cold, cyanotic, or edematous. Additionally, the Centers for Disease Control and Prevention calls out the danger of infection through the misuse of finger sticks. Solutions that address these limitations benefit millions of individuals worldwide. {data['petitioner_name']}'s development of a personalized glucose monitoring system (PGMS) not only addresses these issues but provides personalized blood sugar analysis through the use of artificial intelligence methods and is, therefore, of great merit to the world (Exhibits 1-4, 23, 24). {data['petitioner_name']}'s proposed research on the design of green wireless networks and preventive healthcare systems to create energy-efficient 5G wireless networks that support various healthcare systems, including his future research at {data['current_position_university']} or a similar employer, therefore, has substantial merit (Exhibit 7). Fellow experts in the field provide additional insight into the merit of this endeavor:"
#         )

# # Generate testimonials section
#             testimonials = []
#             for i in range(1, data['num_testimonials'] + 1):
#                     testimonial_content_key = f'testimonial_{i}'
#                     testimonial_author_key = f'testimonial_author_{i}'
#                     testimonial_position_key = f'testimonial_author_position_{i}'
#                     testimonial_organization_key = f'testimonial_author_organization_{i}'

#                     if all(key in data for key in [testimonial_content_key, testimonial_author_key, testimonial_position_key, testimonial_organization_key]):
#                         testimonial_content = data[testimonial_content_key]
#                         testimonial_author = data[testimonial_author_key]
#                         testimonial_position = data[testimonial_position_key]
#                         testimonial_organization = data[testimonial_organization_key]

#                     testimonials.append(
#                         f"• “{testimonial_content}” (Exhibit {i}. {testimonial_author}, {testimonial_position}, {testimonial_organization})"
#                     )

#             self.introduction += "".join(testimonials)

#             f"ii. {data['petitioner_name']}'s proposed endeavor has national importance"
#             f"{data['petitioner_name']}'s proposed endeavor also has broad implications for the United States. In an analysis, the {data['analyzing_organization']} recently characterized the actions and policies of governments relating to {data['field_of_expertise']} as '{data['global_race_description']}'. {data['field_of_expertise']} enables a wide range of advanced applications, from {data['advanced_applications']}. The U.S. has recognized the importance of {data['field_of_expertise']} in the development of a “National Strategy to Secure {data['field_of_expertise']} Implementation Plan” by the U.S. Department of Commerce. Yet, constraints to full deployment remain, including {data['deployment_constraints']}. In a recap of {data['energy_analysis_organization']}'s analysis of {data['energy_analysis_field']} requirements, {data['news_source']} notes key issues, including {data['key_energy_issues']}. These limitations result in {data['transmission_limitations']}. {data['petitioner_name']}'s research, with its innovative use of {data['innovative_techniques']}, directly addresses these limitations (Exhibits 1-4, {data['additional_sources']})."
#             f"{data['petitioner_name']}'s proposed research on {data['proposed_endeavor']}, including their future research at {data['employment_employer']} or a similar employer, is therefore also nationally important (Exhibits [PES-EL])."
#             f"Because {data['petitioner_name']}'s proposed endeavor has both substantial merit and national importance, they satisfy this prong."


#             f"- Compose “Addressing the Multifaceted Assessment” up to 400 words as outlined in Dhanasar, how it effectively demonstrates that {data['petitioner_name']} is well positioned to advance his proposed endeavor. Discuss the key aspects of education, skills, knowledge, and the record of success and emphasize it?"
#             f"IV. {data['petitioner_name']} is well positioned to advance the proposed endeavor of {data['proposed_endeavor']}\n"
#             f"Dhanasar indicates that the second prong of the analysis must consider whether the petitioner is well positioned to advance the proposed endeavor (Dhanasar, at 890). This multifactorial assessment includes an evaluation of the petitioner’s education, skills, knowledge, and record of success in related efforts; a model or plan for future activities; any progress made toward achieving the proposed endeavor; and the interest of potential customers, users, investors, or other relevant entities or individuals (Id.). Importantly, Dhanasar points out the inherent difficulty in 'forecasting feasibility or future success,' even in the presence of a cogent plan and competent execution; therefore, petitioners are not required to show that their proposed endeavor is more likely than not to succeed (Id.) (Exhibit [Dhanasar]).\n"
#             f"Based on this multifactorial assessment, it is clear that {data['petitioner_name']}'s education, experience, expertise, documented record of success, influence in their field, and their future plan have altogether well positioned them to advance the proposed endeavor of {data['proposed_endeavor']}.\n"


#             f"- Compose “Highlighting Education and Record of Success” up to 800 words that should suggest ways to emphasize {data['petitioner_name']}'s education, skills, and record of success in related efforts, particularly focusing on his publications, influential topics, and impact on other researchers. Write in detail how “Highlighting Education and Record of Success” this contributes to a strong case."
#             f"i. Education, Skills, and Knowledge\n"
#             f"{data['petitioner_name']} earned their MS in {data['ms_field']} from {data['ms_university']} and their PhD in {data['phd_field']} from {data['phd_university']}. They are currently {data['current_position']} at {data['current_position_university']} in {data['current_position_location']}. They have published significant research on the application of advanced technologies to issues in {data['field_of_expertise']} (Exhibits [degrees, CV]). Based on this background, {data['petitioner_name']} plans to pursue a position with {data['employment_employer']} or a similar employer, where they will continue their research into {data['proposed_endeavor']} (Exhibit [PES-EL]). Fellow experts have described the importance of {data['petitioner_name']}'s background and experience in more detail in letters of support (Exhibits 1-4).\n"
#             f"ii. Record of Success in Related or Similar Efforts and Interest of Relevant Individuals\n"
#             f"Throughout their time working in the field, {data['petitioner_name']} has built an impressive record of success. As detailed below, their original research on nationally important topics like {data['research_topics']} has been their development of {data['development_description']} (Exhibits 1-[notable citations/notable implementation]). This is an unusually strong record of success for a researcher in {data['field_of_expertise']} and demonstrates {data['petitioner_name']}'s ability to continue pursuing their proposed endeavor of {data['proposed_endeavor']} (Exhibits [notable achievements]).\n"

#             f"- Draft the “Leveraging Testimonials and External Recognition“ up to 800 words from industry leaders and {data['petitioner_name']}'s recognition in the field, leverage the external validation to further strengthen the case of {data['petitioner_name']}. Discuss the specific phrases or points from the testimonials that should be highlighted."
#             f"a) {data['petitioner_name']}'s research has been published in authoritative peer-reviewed journals in their field\n"
#             f"{data['petitioner_name']}'s research has resulted in {data['num_journal_articles']} peer-reviewed journal articles ({data['num_first_authored']} of them first-authored) (Exhibits [publications]). Moreover, these papers have been published in the top journals in {data['field_of_expertise']}, reflecting their peers’ recognition of the value of this research (Exhibits [publications, journal rankings]).\n"
#             f"Experts in the field have submitted letters confirming that {data['petitioner_name']}'s record of successful research has well positioned them to continue advancing the proposed endeavor (Exhibits 1-4).\n"
#             f"b) Researchers from around the world have relied upon {data['petitioner_name']}'s research to further their own investigations in the field\n"
#             f"Not only has {data['petitioner_name']} successfully completed and published the results of their research in the field, but their research has also gone on to influence their peers. That is, {data['petitioner_name']}'s publications have been cited a total of {data['citation_count']} times according to Google Scholar, thereby demonstrating that these publications are widely recognized and relied upon in the field of {data['field_of_expertise']} (Exhibit [citation record]).\n"

#             f"8. **Connecting Research to National Importance:**\n"
#             f"- Draft “Connecting Research to National Importance” up to 800 words. Compose and effectively connect {data['petitioner_name']}'s research on advanced {data['research_subfield']} applications to broader national implications, considering the global competition for technological superiority and the U.S. Department of Commerce's recognition."

#             f"9. **Addressing Potential Concerns:**\n"
#             f"-Draft “Addressing Potential Concerns” up to 700 words. Considering the inherent difficulty in forecasting feasibility, and address this challenge and present a compelling case for {data['petitioner_name']}'s future success in advancing the proposed endeavor. Focus on key points that should be emphasized."

#             f"10. **Creating a Cohesive Narrative:**\n"
#             f"- Lastly, Draft “Creating a Cohesive Narrative” up to 500 words. Ensure that the cover letter presents a cohesive narrative that ties together {data['petitioner_name']}'s education, proposed endeavor, merit, national importance, and his ability to advance the endeavor. Highlight the elements that should be woven seamlessly throughout the letter."

   
# Continue with the rest of your code...

    class CoverLetterData:
        def __init__(self, data):
            self.introduction = (
            f"1. **Introduction to Standard of Proof:**\n"
            f"- Compose an introduction of up to 300 words that elaborates on the 'preponderance of the evidence' standard and how it applies to EB2 NIW petitions, drawing insights from Matter of Dhanasar, Matter of E-M-, and U.S. v. Cardoza-Fonseca"
            f"For Example, Generate like Below"
            f"Note that the standard of proof for petitions filed for National Interest Waiver cases is the 'preponderance of the evidence' standard. See Matter of Dhanasar, 26 I&N Dec. 884, 889 (AAO 2016). Thus, if the petitioner submits relevant, probative, and credible evidence that leads USCIS to believe that the claim is 'more likely than not' or 'probably true,' the petitioner has satisfied the standard of proof. Matter of E-M-, 20 I&N Dec. 77, 79-80 (Comm’r 1989); see also U.S. v. Cardoza-Fonseca, 480 U.S. 421 (1987) (discussing 'more likely than not' as a greater than 50% chance of an occurrence taking place).\n\n"
        )

            self.heading1 = (
            f"- Compose the Highlighting Advanced Degree of up to 500 words that Provide a compelling narrative around {data['petitioner_name']}'s advanced degree, emphasizing the significance of their Ph.D. in {data['degree_field']} from {data['degree_university']}. You have to present this information to showcase {data['petitioner_name']} as a highly qualified professional."
            f"I. {data['petitioner_name']} is a member of the professions holding an advanced degree\n"
            f"I. {data['petitioner_name']} received a PhD in {data['degree_field']} from {data['degree_university']} in {data['degree_year']}. As evidence of this, we are submitting copies of {data['petitioner_name']}'s diploma and transcripts (Exhibit [advanced degree, transcripts]). Since {data['petitioner_name']} completed their education outside the United States, we are also submitting a detailed advisory evaluation of their educational credentials (Exhibit {data['degree_evaluation']}). As such, {data['petitioner_name']} is qualified as a member of the professions holding an advanced degree.\n\n"
        )

            self.heading2 = (
            f"- Compose the Detailing Proposed Endeavor up to 800 words that enhance the description of {data['petitioner_name']}'s proposed endeavor of developing state-of-the-art {data['field_of_expertise']} in the field of {data['research_area']}? All aspects should be emphasized to make it stand out."
            f"II. {data['petitioner_name']}'s proposed endeavor is in the field of {data['field_of_expertise']} to {data['proposed_endeavor_description']}\n"
            f"As an expert in the field of {data['field_of_expertise']}, {data['petitioner_name']}'s proposed endeavor is to {data['proposed_endeavor']} (Exhibits [PES-EL]).\n"
            f"This petition waives the job offer requirement, and the petitioner's proposed endeavor is separate from their proposed employment. However, we are submitting {data['petitioner_name']}'s plans for employment in the field to confirm their commitment and capacity to advance their proposed endeavor. Based on their education and research background, {data['petitioner_name']} plans to be employed as a {data['employment_position']} at {data['employment_employer']} or a similar employer (Exhibit [PES-EL]). {data['petitioner_name']} intends to continue their research on {data['proposed_endeavor']} (Exhibits [PES-EL]). This said, the focus of this prong should be on the proposed endeavor itself rather than {data['petitioner_name']}'s employment.\n\n"
        )

            self.heading3 = (
            f"- Compose the “Emphasizing Merit and National Importance” up to 800 words that show Given {data['petitioner_name']}'s research applications and the testimonials provided, you have to strengthen the argument for the substantial merit and national importance of his proposed endeavor. Discuss specific examples or points that should be highlighted."
            f"{data['petitioner_name']}'s proposed endeavor of {data['proposed_endeavor']} has both substantial merit and national importance"
            f"{data['petitioner_name']}'s research related to {data['proposed_endeavor']} has great substantial merit and national importance. Among other applications, their research is relevant to {data['research_applications']} (Exhibits 1-4, [publications, testimonial letters])."
            f"i. {data['petitioner_name']}'s proposed endeavor has substantial merit"
            f"{data['petitioner_name']}'s research advancing their proposed endeavor is of great importance because it allows for {data['benefits']} (Exhibits 1-4, {data['sources']}). Fellow experts in the field provide additional insight into the merit of this endeavor:"
            f"{data['petitioner_name']}'s research advancing his proposed endeavor is of great importance because it allows for improved healthcare to global communities. It is estimated that there are currently {data['diabetes_population']} individuals afflicted with diabetes globally, or {data['diabetes_percentage']} of the world’s population. That number is projected to grow to approximately {data['projected_diabetes_population']} by {data['projected_year']}. Health officials encourage diabetics to monitor their blood sugar daily. Finger sticks, in which the individual pricks his or her fingertip to provide blood for analysis by a portable sensor, remains the primary method for this test. However, finger sticks have several limitations, including the pain of repeated jabs to fingertips, as well as difficulty of application on individuals whose hands are swollen, cold, cyanotic, or edematous. Additionally, the Centers for Disease Control and Prevention calls out the danger of infection through the misuse of finger sticks. Solutions that address these limitations benefit millions of individuals worldwide. {data['petitioner_name']}'s development of a personalized glucose monitoring system (PGMS) not only addresses these issues but provides personalized blood sugar analysis through the use of artificial intelligence methods and is, therefore, of great merit to the world (Exhibits 1-4, 23, 24). {data['petitioner_name']}'s proposed research on the design of green wireless networks and preventive healthcare systems to create energy-efficient 5G wireless networks that support various healthcare systems, including his future research at {data['current_position_university']} or a similar employer, therefore, has substantial merit (Exhibit 7). Fellow experts in the field provide additional insight into the merit of this endeavor:"
        )

            testimonials = []
            for testimonial in data['testimonials']:
                testimonials.append(
                    f"• “{testimonial['content']}” (Exhibit {len(testimonials) + 1}. {testimonial['author']}, "
                    f"{testimonial['position']}, {testimonial['organization']})"
                )
            
            self.heading3 += "".join(testimonials)

            self.heading3 += (
            f"ii. {data['petitioner_name']}'s proposed endeavor has national importance"
            f"{data['petitioner_name']}'s proposed endeavor also has broad implications for the United States. In an analysis, the {data['analyzing_organization']} recently characterized the actions and policies of governments relating to {data['field_of_expertise']} as '{data['global_race_description']}'. {data['field_of_expertise']} enables a wide range of advanced applications, from {data['advanced_applications']}. The U.S. has recognized the importance of {data['field_of_expertise']} in the development of a “National Strategy to Secure {data['field_of_expertise']} Implementation Plan” by the U.S. Department of Commerce. Yet, constraints to full deployment remain, including {data['deployment_constraints']}. In a recap of {data['energy_analysis_organization']}'s analysis of {data['energy_analysis_field']} requirements, {data['news_source']} notes key issues, including {data['key_energy_issues']}. These limitations result in {data['transmission_limitations']}. {data['petitioner_name']}'s research, with its innovative use of {data['innovative_techniques']}, directly addresses these limitations (Exhibits 1-4, {data['additional_sources']})."
            f"{data['petitioner_name']}'s proposed research on {data['proposed_endeavor']}, including their future research at {data['employment_employer']} or a similar employer, is therefore also nationally important (Exhibits [PES-EL])."
            f"Because {data['petitioner_name']}'s proposed endeavor has both substantial merit and national importance, they satisfy this prong."
        )

            self.heading4 = (
            f"- Compose “Addressing the Multifaceted Assessment” up to 400 words as outlined in Dhanasar, how it effectively demonstrates that {data['petitioner_name']} is well positioned to advance his proposed endeavor. Discuss the key aspects of education, skills, knowledge, and the record of success and emphasize it?"
            f"IV. {data['petitioner_name']} is well positioned to advance the proposed endeavor of {data['proposed_endeavor']}\n"
            f"Dhanasar indicates that the second prong of the analysis must consider whether the petitioner is well positioned to advance the proposed endeavor (Dhanasar, at 890). This multifactorial assessment includes an evaluation of the petitioner’s education, skills, knowledge, and record of success in related efforts; a model or plan for future activities; any progress made toward achieving the proposed endeavor; and the interest of potential customers, users, investors, or other relevant entities or individuals (Id.). Importantly, Dhanasar points out the inherent difficulty in 'forecasting feasibility or future success,' even in the presence of a cogent plan and competent execution; therefore, petitioners are not required to show that their proposed endeavor is more likely than not to succeed (Id.) (Exhibit [Dhanasar]).\n"
            f"Based on this multifactorial assessment, it is clear that {data['petitioner_name']}'s education, experience, expertise, documented record of success, influence in their field, and their future plan have altogether well positioned them to advance the proposed endeavor of {data['proposed_endeavor']}.\n"
        )

            self.heading5 = (
            f"- Compose “Highlighting Education and Record of Success” up to 800 words that should suggest ways to emphasize {data['petitioner_name']}'s education, skills, and record of success in related efforts, particularly focusing on his publications, influential topics, and impact on other researchers. Write in detail how “Highlighting Education and Record of Success” this contributes to a strong case."
            f"i. Education, Skills, and Knowledge\n"
            f"{data['petitioner_name']} earned their MS in {data['ms_field']} from {data['ms_university']} and their PhD in {data['phd_field']} from {data['phd_university']}. They are currently {data['current_position']} at {data['current_position_university']} in {data['current_position_location']}. They have published significant research on the application of advanced technologies to issues in {data['field_of_expertise']} (Exhibits [degrees, CV]). Based on this background, {data['petitioner_name']} plans to pursue a position with {data['employment_employer']} or a similar employer, where they will continue their research into {data['proposed_endeavor']} (Exhibit [PES-EL]). Fellow experts have described the importance of {data['petitioner_name']}'s background and experience in more detail in letters of support (Exhibits 1-4).\n"
            f"ii. Record of Success in Related or Similar Efforts and Interest of Relevant Individuals\n"
            f"Throughout their time working in the field, {data['petitioner_name']} has built an impressive record of success. As detailed below, their original research on nationally important topics like {data['research_topics']} has been their development of {data['development_description']} (Exhibits 1-[notable citations/notable implementation]). This is an unusually strong record of success for a researcher in {data['field_of_expertise']} and demonstrates {data['petitioner_name']}'s ability to continue pursuing their proposed endeavor of {data['proposed_endeavor']} (Exhibits [notable achievements]).\n"
        )

            self.heading6 = (
            f"- Draft the “Leveraging Testimonials and External Recognition“ up to 800 words from industry leaders and {data['petitioner_name']}'s recognition in the field, leverage the external validation to further strengthen the case of {data['petitioner_name']}. Discuss the specific phrases or points from the testimonials that should be highlighted."
            f"a) {data['petitioner_name']}'s research has been published in authoritative peer-reviewed journals in their field\n"
            f"{data['petitioner_name']}'s research has resulted in {data['num_journal_articles']} peer-reviewed journal articles ({data['num_first_authored']} of them first-authored) (Exhibits [publications]). Moreover, these papers have been published in the top journals in {data['field_of_expertise']}, reflecting their peers’ recognition of the value of this research (Exhibits [publications, journal rankings]).\n"
            f"Experts in the field have submitted letters confirming that {data['petitioner_name']}'s record of successful research has well positioned them to continue advancing the proposed endeavor (Exhibits 1-4).\n"
            f"b) Researchers from around the world have relied upon {data['petitioner_name']}'s research to further their own investigations in the field\n"
            f"Not only has {data['petitioner_name']} successfully completed and published the results of their research in the field, but their research has also gone on to influence their peers. That is, {data['petitioner_name']}'s publications have been cited a total of {data['citation_count']} times according to Google Scholar, thereby demonstrating that these publications are widely recognized and relied upon in the field of {data['field_of_expertise']} (Exhibit [citation record]).\n"
        )

            self.heading7 = (
    f"8. **Connecting Research to National Importance:**\n"
    f"- Draft “Connecting Research to National Importance” up to 800 words. Compose and effectively connect {data['petitioner_name']}'s research on advanced {data['research_subfield']} applications to broader national implications, considering the global competition for technological superiority and the U.S. Department of Commerce's recognition."
)

            self.heading8 = (
    f"9. **Addressing Potential Concerns:**\n"
    f"-Draft “Addressing Potential Concerns” up to 700 words. Considering the inherent difficulty in forecasting feasibility, and address this challenge and present a compelling case for {data['petitioner_name']}'s future success in advancing the proposed endeavor. Focus on key points that should be emphasized."
)

            self.heading9 = (
    f"10. **Creating a Cohesive Narrative:**\n"
    f"- Lastly, Draft “Creating a Cohesive Narrative” up to 500 words. Ensure that the cover letter presents a cohesive narrative that ties together {data['petitioner_name']}'s education, proposed endeavor, merit, national importance, and his ability to advance the endeavor. Highlight the elements that should be woven seamlessly throughout the letter."
)


    prompts = CoverLetterData(data)

    # Dictionary to store generated responses
    responses = {}

    # Loop through each attribute in the CoverLetterData class and generate responses
    for attribute_name, attribute_value in vars(prompts).items():
        if isinstance(attribute_value, str):  # Check if the attribute is a string
            response = await generate_response(attribute_value)
            responses[attribute_name] = response

    # Load the Word document template (replace with your template file path)
    doc = DocxTemplate('generated-document.docx')

    # Define the context with variables for template replacement
    context = {
    'petitioner_name': data['petitioner_name'],

    # Prompts
    "introduction": responses['introduction'],
    

}

    # Render the template with the context
    # ...

# Render the template with the context
    doc.render(context)

    # Provide a BytesIO object to hold the generated document
# Save the document
    doc.save('generated-document-testing.docx')

# Provide a BytesIO object to hold the generated document
    with open('generated-document-testing.docx', 'rb') as file:
        doc_buffer = file.read()

    return doc_buffer

def main(data):

    st.title("Generate Word Document")

    
# Create a new dictionary to store modified values
    updated_data = {}

# Input fields for data (excluding testimonials)
    for key, value in data.items():
        if key != 'testimonials':
            label = key.replace('_', ' ').title()
            input_value = st.text_input(label, value=value, key=key, placeholder=value)
            updated_data[key] = input_value

# Input field for num_testimonials
    num_testimonials = st.number_input("Number of Testimonials", min_value=1, value=data.get('num_testimonials', 1))

# Input fields for testimonials
    testimonials_list = []
    for i in range(1, num_testimonials + 1):
        st.subheader(f'Testimonial {i}')
        testimonial_dict = {}
        for testimonial_key, testimonial_value in data['testimonials'][0].items():  # Assuming all testimonials have the same structure
            testimonial_label = testimonial_key.replace('_', ' ').title()
            testimonial_input_key = f'testimonial_{i}_{testimonial_key}'
            testimonial_input_value = st.text_input(testimonial_label, key=testimonial_input_key, value=testimonial_value)
            testimonial_dict[testimonial_key] = testimonial_input_value
        testimonials_list.append(testimonial_dict)

    updated_data['testimonials'] = testimonials_list

    if st.button("Generate Document"):
        if all(value for value in data.values()):
            st.write("Generating document...")
            doc_buffer = asyncio.run(generate_document(data))

            # Convert the Document object to bytes
            # doc_bytes = doc_buffer.read()

            # Generate a unique filename using a timestamp
            timestamp = time.strftime("%Y%m%d%H%M%S")
            filename = f'generated-document-{timestamp}.docx'

            # Provide a download link for users
            st.markdown(get_binary_file_downloader_html(doc_buffer, filename, 'Download Document'), unsafe_allow_html=True)

# Rest of the code remains the same

# Function to create a download link
def get_binary_file_downloader_html(bin_file, filename, label):
    b64 = base64.b64encode(bin_file).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{label}</a>'
    return href

if __name__ == '__main__':
    main(data)
