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
    "Name of the Petitioner": "Mudasar Latif Memon",
    "Highest Qualification": "Ph.D.",
    "Main Field/Industry": "Electronic and Electrical Engineering",
    "University Name of Highest Degree": "Sungkyunkwan University, South Korea",
    "Date of Graduation": "25 August 2019",
    "Subfields/Specializations": [
        "Artificial Intelligence for Energy Efficiency in 5G Wireless Networks",
        "Green and Emerging Wireless Networks",
        "Intelligent Solutions for Preventive Healthcare Systems",
    ],
    "Top Skills": [
        "AI-based Energy Saving in 5G Wireless Networks",
        "Backscatter Communication for future Battery-free Communication in IoTs",
        "AI-based Algorithm Designing for DRX in 5G",
        "AI-based Non-invasive Blood Glucose Monitoring for Preventive Healthcare System",
        "AI-based Activity and Food Recognition System for Preventive Healthcare System",
    ],
    "Areas in the Field": {
        "Area-1": {
            "Area Name": "Artificial Intelligence based Energy savings for 5G Network connected cellphones",
            "Unique Contribution": "Proposed AI-based flexible Discontinuous Reception (DRX) for 5G network.",
            "Unique Skillset": "Used Long Short-term Memory (LSTM) Networks in DRX of 5G wireless networks.",
        },
        "Area-2": {
            "Area Name": "Backscatter Communication for battery-free Internet of Things",
            "Unique Contribution": "Detailed and comparative study on Backscatter Communication.",
            "Unique Skillset": "Summarized perspectives of Backscatter Communication for battery-free IoT devices.",
        },
        "Area-3": {
            "Area Name": "Artificial Intelligence for Preventive Healthcare Systems",
            "Unique Contribution": "Presented various AI-based solutions for preventive healthcare systems.",
            "Unique Skillset": "Used machine learning models for a personalized glucose monitoring system and AI-based activity and food recognition system.",
        },
    },
    "Top Achievements in Field/Industry": [
        "Proposed AI-based DRX mechanism for multiple beam communications in 5G networks.",
        "Suggested AI-DRX algorithm using a ten-state model for dynamic short or long sleep cycles.",
        "Achieved energy efficiency of 59% on trace 1 and 95% on trace 2 with AI-DRX.",
        "Developed Backscatter Communication as a potential technology for future battery-free communication in small IoT devices.",
        "Developed a Personalized Glucose Monitoring System using AdaBoost algorithm.",
        "Proposed an automated system for monitoring activities and food types.",
    ],
    "Plans for Continued Engagement in the Field": {
        "Proposed Endeavor": [
            "Build on extensive experience with AI in Electronic and Electrical Engineering.",
            "Design state-of-the-art Green wireless networks for future wireless networks and healthcare systems.",
            "Circulate work through peer-reviewed publications and conference presentations.",
        ],
        "Evidence of Employment/Employability in the Field": [
            "Applied for a post-doctoral researcher position in the same area.",
            "May work with Prof. Navrati Saxena at San Jose State University.",
        ],
        "How Past Work has Prepared for Proposed Endeavor": [
            "Use knowledge gained from AI in 5G networks, Backscatter Communication, and Machine Learning for personalized glucose monitoring.",
            "Skills include applied AI models, Backscatter Communication, AI for Healthcare systems, algorithm design, technical writing, and familiarity with Matlab, Python, and Latex.",
        ],
    },
    "Media Coverage": [
        {
            "Coverage 1": {
                "Media Outlet": "Korea Herald",
                "Link": "http://www.koreaherald.com/view.php?ud=20180808000636",
                "Focus": 'The second paragraph of the news focuses on my research on Backscatter Communication, specifically "ways to convert small amounts of energy from wireless signals back into electrical energy."',
            },
        },
    ],
    "Awards and Scholarships": [
        {
            "Research Award": {
                "Award": "1st Prize in Superior Research Award",
                "Recipient": "Mudasar Latif Memon",
                "Link": "https://skb.skku.edu/eng_ice/intro/notice02.do?mode=view&link=http://app.skku.edu/emate_app/bbs/b1805133145.nsf/api/data/documents/unid/E5D6A67B6600BFCF4925844800090DF1?multipart=false&article.offset=130&articleLimit=10",
                "Awarding Institution": "College of Information and Communication Engineering (CICE), Sungkyunkwan University",
                "Eligibility": "All graduating Ph.D. Students of CICE",
                "Number of Competitors/Winners": "2",
                "Selection Criteria": "Number of research publications in the last 6 months",
                "Judges": "Professors and Dean, College of Information and Communication Engineering",
            },
        },
        {
            "Ph.D. Scholarship": {
                "Award": "Foreign Funded Ph.D. Scholarship",
                "Recipient": "Mudasar Latif Memon",
                "Link": "https://www.hec.gov.pk/english/scholarshipsgrants/UESTPS-UETS/Pages/Introduction.aspx",
                "Awarding Institution": "Higher Education Commission (HEC), Pakistan",
                "Eligibility": "All Pakistani Nationals having graduation",
                "Number of Competitors/Winners": "Approximately 30 per batch to different foreign countries",
                "Selection Criteria": "National Level Competition based on Academic record, personal profile, and Graduate Record Examination Score",
                "Judges": "Higher Education Commission, Pakistan/Government Officials",
            },
        },
        {
            "Community College Administrator Program": {
                "Award": "Community College Administrator Program (CCAP) at UMass, MA, USA",
                "Recipient": "Mudasar Latif Memon",
                "Link": "https://www.umass.edu/cie/our-work/professional-development-community-college-administrators-pakistan-2015-2017",
                "Awarding Institution": "U.S. Department of State’s Bureau of Educational and Cultural Affairs",
                "Eligibility": "All nominated administrators of community colleges/Technical Colleges/Technical Universities of Pakistan",
                "Number of Competitors/Winners": "Approximately 20 per batch",
                "Selection Criteria": "United States Education Foundation Pakistan (USEFP) Team, US state department Team",
                "Judges": "Officials from U.S. Department of State’s Bureau of Educational and Cultural Affairs",
            },
        },
    ],
    "Judging Experience": [
        {
            "Ph.D. Defense Judge": {
                "Date": "April 19, 2021",
                "Institution": "Quaid-e-Awam University of Engineering and Technology, Nawabshah, Sindh, Pakistan",
                "Review": "One review",
            },
        },
        {
            "Masters Defense Judge": {
                "Date": "March 17, 2021",
                "Institution": "Quaid-e-Awam University of Engineering and Technology, Nawabshah, Sindh, Pakistan",
                "Review": "One review",
            },
        },
        {
            "Board of Studies Meeting": {
                "Meeting Type": "4th Board of Studies (BoS) Meeting",
                "Institution": "Department of Telecommunication, Quaid-e-Awam University of Engineering and Technology, Nawabshah, Sindh, Pakistan",
                "Date": "[Not specified]",
            },
        },
    ],
    "Leadership Roles": [
        {
            "Committee Membership": {
                "Committee Name": "IEEE Consumer Communications and Networking Conference (CCNC) Technical Program Committee",
                "Dates of Service": "Nov 16, 2020",
                "Role": "Technical Program Committee Member",
                "Organization's Prestige": "[Link to CCNC 2021](https://wcnc2020.ieee-wcnc.org/)",
            },
        },
        {
            "Committee Membership": {
                "Committee Name": "IEEE Wireless Communications and Networking Conference",
                "Dates of Service": "25-28 May 2020 (Virtual Conference)",
                "Role": "Technical Program Committee Member",
                "Organization's Prestige": "[Link to WCNC 2020](https://wcnc2020.ieee-wcnc.org/)",
            },
        },
    ],
    "Journal Articles": [
        {
            "Article 1": {
                "Authors": "Mudasar Latif Memon, Mukesh Kumar Maheshwari, Navrati Saxena, Abhishek Roy, and Dong Ryeol Shin",
                "Title": "Artificial Intelligence-Based Discontinuous Reception for Energy Saving in 5G Networks",
                "Journal": "Electronics",
                "Volume": "8, Number: 7",
                "Year": "2019",
                "Page": "778",
                "Impact Factor": "2.4 SCI-E",
                "Availability": "[Available Online]",
            },
        },
        {
            "Article 2": {
                "Authors": "Mudasar Latif Memon, Mukesh Kumar Maheshwari, Dong Ryeol Shin, Abhishek Roy, and Navrati Saxena",
                "Title": "Deep‐DRX: A framework for deep learning–based discontinuous reception in 5G wireless networks",
                "Journal": "Transactions on Emerging Telecommunications Technologies",
                "Volume": "30, Number: 3",
                "Year": "2019",
                "Page": "e3579",
                "Impact Factor": "1.6 SCI-E",
                "Availability": "[Available Online]",
            },
        },
        {
            "Article 3": {
                "Authors": "Mudasar Latif Memon, Navrati Saxena, Abhishek Roy, and Dong Ryeol Shin",
                "Title": "Backscatter communications: Inception of the battery-free era—A comprehensive survey",
                "Journal": "Electronics",
                "Volume": "8, Number: 2",
                "Year": "2019",
                "Page": "129",
                "Impact Factor": "2.4 SCI-E",
                "Availability": "[Available Online]",
            },
        },
    ],
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
            {"role": "system", "content": f"You are US Immigration attorney assistant having expertise in writing I-140 petition cover letter, presenting strong arguments for Employement Based 2nd preference (EB2) National Interest Waiver (NIW) petitions. Your role is to create and compose a comprehensive cover letter for an EB-2 National Interest Waiver petition, to be written by the petitioner, based on the provided information using instructions given for user content , use data variables in it"},
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

    class CoverLetterData:
        def __init__(self, data):

            self.introduction = (
            f"1. **Introduction to Standard of Proof:**\n"
            f"- Compose an introduction of up to 300 words that elaborates on the 'preponderance of the evidence' standard and how it applies to EB2 NIW petitions, drawing insights from Matter of Dhanasar, Matter of E-M-, and U.S. v. Cardoza-Fonseca. "
            f"For example, generate like below:\n"
            f"Note that the standard of proof for petitions filed for National Interest Waiver cases is the 'preponderance of the evidence' standard. "
            f"See Matter of Dhanasar, 26 I&N Dec. 884, 889 (AAO 2016). Thus, if the petitioner submits relevant, probative, and credible evidence "
            f"that leads USCIS to believe that the claim is 'more likely than not' or 'probably true,' the petitioner has satisfied the standard of proof. "
            f"Matter of E-M-, 20 I&N Dec. 77, 79-80 (Comm’r 1989); see also U.S. v. Cardoza-Fonseca, 480 U.S. 421 (1987) (discussing 'more likely than not' "
            f"as a greater than 50% chance of an occurrence taking place).\n\n"
        )
            
            self.doctor_title = data.get('doctor_title', '')
            self.field_of_study = data.get('field_of_study', '')
            self.university = data.get('university', '')
            self.graduation_year = data.get('graduation_year', '')
            self.gender = data.get('gender', '')

            def generate_content(self):
                content_template = "{doctor_title} received a PhD in {field_of_study} from {university} in {graduation_year}. " \
                           "As evidence of this, we are submitting copies of {doctor_title}'s diploma and transcripts (Exhibit 6). " \
                           "Since {doctor_title} completed {his_or_her} education outside the United States, " \
                           "we are also submitting a detailed advisory evaluation of {his_or_her} educational credentials (Exhibit 6). " \
                           "As such, {doctor_title} is qualified as a member of the professions holding an advanced degree."

                return content_template.format(
                    doctor_title=self.doctor_title,
                    field_of_study=self.field_of_study,
                    university=self.university,
                    graduation_year=self.graduation_year,
                    his_or_her="his" if self.gender == 'male' else "her"
        )


#             self.heading3 = (
#             f"- Compose the “Emphasizing Merit and National Importance” up to 800 words that show Given {data['petitioner_name']}'s research applications and the testimonials provided, you have to strengthen the argument for the substantial merit and national importance of his proposed endeavor. Discuss specific examples or points that should be highlighted."
#             f"{data['petitioner_name']}'s proposed endeavor of {data['proposed_endeavor']} has both substantial merit and national importance"
#             f"{data['petitioner_name']}'s research related to {data['proposed_endeavor']} has great substantial merit and national importance. Among other applications, their research is relevant to {data['research_applications']} (Exhibits 1-4, [publications, testimonial letters])."
#             f"i. {data['petitioner_name']}'s proposed endeavor has substantial merit"
#             f"{data['petitioner_name']}'s research advancing their proposed endeavor is of great importance because it allows for {data['benefits']} (Exhibits 1-4, {data['sources']}). Fellow experts in the field provide additional insight into the merit of this endeavor:"
#             f"{data['petitioner_name']}'s research advancing his proposed endeavor is of great importance because it allows for improved healthcare to global communities. It is estimated that there are currently {data['diabetes_population']} individuals afflicted with diabetes globally, or {data['diabetes_percentage']} of the world’s population. That number is projected to grow to approximately {data['projected_diabetes_population']} by {data['projected_year']}. Health officials encourage diabetics to monitor their blood sugar daily. Finger sticks, in which the individual pricks his or her fingertip to provide blood for analysis by a portable sensor, remains the primary method for this test. However, finger sticks have several limitations, including the pain of repeated jabs to fingertips, as well as difficulty of application on individuals whose hands are swollen, cold, cyanotic, or edematous. Additionally, the Centers for Disease Control and Prevention calls out the danger of infection through the misuse of finger sticks. Solutions that address these limitations benefit millions of individuals worldwide. {data['petitioner_name']}'s development of a personalized glucose monitoring system (PGMS) not only addresses these issues but provides personalized blood sugar analysis through the use of artificial intelligence methods and is, therefore, of great merit to the world (Exhibits 1-4, 23, 24). {data['petitioner_name']}'s proposed research on the design of green wireless networks and preventive healthcare systems to create energy-efficient 5G wireless networks that support various healthcare systems, including his future research at {data['current_position_university']} or a similar employer, therefore, has substantial merit (Exhibit 7). Fellow experts in the field provide additional insight into the merit of this endeavor:"
#         )

#             testimonials = []
#             for i in range(1, int(data['num_testimonials']) + 1):  # Convert num_testimonials to int
#                 testimonials.append(
#                 f"• “{data[f'testimonial_{i}']}” (Exhibit {i + 1}. {data[f'testimonial_author_{i}']}, "
#                 f"{data[f'testimonial_author_position_{i}']}, {data[f'testimonial_author_organization_{i}']})"
#             )

#             self.heading3 += "".join(testimonials)

#             self.heading3 += (
#             f"ii. {data['petitioner_name']}'s proposed endeavor has national importance"
#             f"{data['petitioner_name']}'s proposed endeavor also has broad implications for the United States. In an analysis, the {data['analyzing_organization']} recently characterized the actions and policies of governments relating to {data['field_of_expertise']} as '{data['global_race_description']}'. {data['field_of_expertise']} enables a wide range of advanced applications, from {data['advanced_applications']}. The U.S. has recognized the importance of {data['field_of_expertise']} in the development of a “National Strategy to Secure {data['field_of_expertise']} Implementation Plan” by the U.S. Department of Commerce. Yet, constraints to full deployment remain, including {data['deployment_constraints']}. In a recap of {data['energy_analysis_organization']}'s analysis of {data['energy_analysis_field']} requirements, {data['news_source']} notes key issues, including {data['key_energy_issues']}. These limitations result in {data['transmission_limitations']}. {data['petitioner_name']}'s research, with its innovative use of {data['innovative_techniques']}, directly addresses these limitations (Exhibits 1-4, {data['additional_sources']})."
#             f"{data['petitioner_name']}'s proposed research on {data['proposed_endeavor']}, including their future research at {data['employment_employer']} or a similar employer, is therefore also nationally important (Exhibits [PES-EL])."
#             f"Because {data['petitioner_name']}'s proposed endeavor has both substantial merit and national importance, they satisfy this prong."
#         )

#             self.heading4 = (
#             f"- Compose “Addressing the Multifaceted Assessment” up to 400 words as outlined in Dhanasar, how it effectively demonstrates that {data['petitioner_name']} is well positioned to advance his proposed endeavor. Discuss the key aspects of education, skills, knowledge, and the record of success and emphasize it?"
#             f"IV. {data['petitioner_name']} is well positioned to advance the proposed endeavor of {data['proposed_endeavor']}\n"
#             f"Dhanasar indicates that the second prong of the analysis must consider whether the petitioner is well positioned to advance the proposed endeavor (Dhanasar, at 890). This multifactorial assessment includes an evaluation of the petitioner’s education, skills, knowledge, and record of success in related efforts; a model or plan for future activities; any progress made toward achieving the proposed endeavor; and the interest of potential customers, users, investors, or other relevant entities or individuals (Id.). Importantly, Dhanasar points out the inherent difficulty in 'forecasting feasibility or future success,' even in the presence of a cogent plan and competent execution; therefore, petitioners are not required to show that their proposed endeavor is more likely than not to succeed (Id.) (Exhibit [Dhanasar]).\n"
#             f"Based on this multifactorial assessment, it is clear that {data['petitioner_name']}'s education, experience, expertise, documented record of success, influence in their field, and their future plan have altogether well positioned them to advance the proposed endeavor of {data['proposed_endeavor']}.\n"
#         )

#             self.heading5 = (
#             f"- Compose “Highlighting Education and Record of Success” up to 800 words that should suggest ways to emphasize {data['petitioner_name']}'s education, skills, and record of success in related efforts, particularly focusing on his publications, influential topics, and impact on other researchers. Write in detail how “Highlighting Education and Record of Success” this contributes to a strong case."
#             f"i. Education, Skills, and Knowledge\n"
#             f"{data['petitioner_name']} earned their MS in {data['ms_field']} from {data['ms_university']} and their PhD in {data['phd_field']} from {data['phd_university']}. They are currently {data['current_position']} at {data['current_position_university']} in {data['current_position_location']}. They have published significant research on the application of advanced technologies to issues in {data['field_of_expertise']} (Exhibits [degrees, CV]). Based on this background, {data['petitioner_name']} plans to pursue a position with {data['employment_employer']} or a similar employer, where they will continue their research into {data['proposed_endeavor']} (Exhibit [PES-EL]). Fellow experts have described the importance of {data['petitioner_name']}'s background and experience in more detail in letters of support (Exhibits 1-4).\n"
#             f"ii. Record of Success in Related or Similar Efforts and Interest of Relevant Individuals\n"
#             f"Throughout their time working in the field, {data['petitioner_name']} has built an impressive record of success. As detailed below, their original research on nationally important topics like {data['research_topics']} has been their development of {data['development_description']} (Exhibits 1-[notable citations/notable implementation]). This is an unusually strong record of success for a researcher in {data['field_of_expertise']} and demonstrates {data['petitioner_name']}'s ability to continue pursuing their proposed endeavor of {data['proposed_endeavor']} (Exhibits [notable achievements]).\n"
#         )

#             self.heading6 = (
#             f"- Draft the “Leveraging Testimonials and External Recognition“ up to 800 words from industry leaders and {data['petitioner_name']}'s recognition in the field, leverage the external validation to further strengthen the case of {data['petitioner_name']}. Discuss the specific phrases or points from the testimonials that should be highlighted."
#             f"a) {data['petitioner_name']}'s research has been published in authoritative peer-reviewed journals in their field\n"
#             f"{data['petitioner_name']}'s research has resulted in {data['num_journal_articles']} peer-reviewed journal articles ({data['num_first_authored']} of them first-authored) (Exhibits [publications]). Moreover, these papers have been published in the top journals in {data['field_of_expertise']}, reflecting their peers’ recognition of the value of this research (Exhibits [publications, journal rankings]).\n"
#             f"Experts in the field have submitted letters confirming that {data['petitioner_name']}'s record of successful research has well positioned them to continue advancing the proposed endeavor (Exhibits 1-4).\n"
#             f"b) Researchers from around the world have relied upon {data['petitioner_name']}'s research to further their own investigations in the field\n"
#             f"Not only has {data['petitioner_name']} successfully completed and published the results of their research in the field, but their research has also gone on to influence their peers. That is, {data['petitioner_name']}'s publications have been cited a total of {data['citation_count']} times according to Google Scholar, thereby demonstrating that these publications are widely recognized and relied upon in the field of {data['field_of_expertise']} (Exhibit [citation record]).\n"
#         )

#             self.heading7 = (
#     f"8. **Connecting Research to National Importance:**\n"
#     f"- Draft “Connecting Research to National Importance” up to 800 words. Compose and effectively connect {data['petitioner_name']}'s research on advanced {data['research_subfield']} applications to broader national implications, considering the global competition for technological superiority and the U.S. Department of Commerce's recognition."
# )

#             self.heading8 = (
#     f"9. **Addressing Potential Concerns:**\n"
#     f"-Draft “Addressing Potential Concerns” up to 700 words. Considering the inherent difficulty in forecasting feasibility, and address this challenge and present a compelling case for {data['petitioner_name']}'s future success in advancing the proposed endeavor. Focus on key points that should be emphasized."
# )

#             self.heading9 = (
#     f"10. **Creating a Cohesive Narrative:**\n"
#     f"- Lastly, Draft “Creating a Cohesive Narrative” up to 500 words. Ensure that the cover letter presents a cohesive narrative that ties together {data['petitioner_name']}'s education, proposed endeavor, merit, national importance, and his ability to advance the endeavor. Highlight the elements that should be woven seamlessly throughout the letter."
# )


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
    'name': data['Name of the Petitioner'],

    # Prompts
    "intro": responses['intro'],
    "heading1": responses['heading1'],
    "heading2": responses['heading2'],
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
    
    def process_nested_data(data, prefix=""):
        updated_data = {}
        for key, value in data.items():
            if isinstance(value, dict):
                st.subheader(key)
                updated_data.update(process_nested_data(value, prefix=f"{prefix}{key}_"))
            elif isinstance(value, list):
                st.subheader(key)
                for i, item in enumerate(value, start=1):
                    if isinstance(item, dict):
                        updated_data.update(process_nested_data(item, prefix=f"{prefix}{key}_{i}_"))
                    else:
                        label = key.replace('_', ' ').title()
                        input_key = f'{prefix}{key}_{i}'
                        input_value = st.text_input(label, value=item, key=input_key, placeholder=item)
                        updated_data[input_key] = input_value
            else:
                label = key.replace('_', ' ').title()
                input_key = f'{prefix}{key}'
                input_value = st.text_input(label, value=value, key=input_key, placeholder=value)
                updated_data[input_key] = input_value
        return updated_data
    

    
    # Input fields for data
    updated_data = process_nested_data(data)


    if st.button("Generate Document"):
        if all(value for value in updated_data.values()):
            st.write("Generating document...")
            doc_buffer = asyncio.run(generate_document(updated_data))

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
