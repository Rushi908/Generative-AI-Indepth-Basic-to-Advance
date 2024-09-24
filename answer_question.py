import pandas as pd
import openai
from fpdf import FPDF

# Set up your OpenAI API key
openai.api_key = 'your-openai-api-key'  # Replace with your actual OpenAI API key

# Load the Excel file
file_path = 'questions.xlsx'  # Update with your file path
sheet_name = 'Sheet1'  # Update if your sheet has a different name
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Assuming the Excel file has a column named "Questions"
questions_column = 'Questions'  # Update with the actual column name containing the questions

# Custom prompt to wrap around the question
def create_prompt(question):
    return f"You are an AI assistant. Please provide a detailed, clear, and concise answer to the following question:\n\nQuestion: {question}\n\nAnswer:"

# Function to get a response from the LLM model using the custom prompt
def get_llm_answer(question):
    prompt = create_prompt(question)
    try:
        response = openai.Completion.create(
            engine="text-davinci-003",  # Use GPT-3 or GPT-4 depending on your API access
            prompt=prompt,
            max_tokens=150,
            temperature=0.7
        )
        return response['choices'][0]['text'].strip()
    except Exception as e:
        return f"Error: {str(e)}"

# Create a PDF object
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()

# Set title and font for the PDF
pdf.set_font('Arial', 'B', 16)
pdf.cell(200, 10, txt="Questions and Answers", ln=True, align='C')
pdf.ln(10)  # Line break

# Iterate through each question, get the LLM answer, and add it to the PDF
pdf.set_font('Arial', '', 12)
for index, question in df[questions_column].iteritems():
    pdf.set_font('Arial', 'B', 12)  # Bold for questions
    pdf.multi_cell(0, 10, f"Question {index + 1}: {question}")
    
    answer = get_llm_answer(question)
    
    pdf.set_font('Arial', '', 12)  # Regular font for answers
    pdf.multi_cell(0, 10, f"Answer: {answer}")
    pdf.ln(5)  # Add some space between questions

# Save the PDF to a file
output_pdf_file = 'questions_and_answers.pdf'
pdf.output(output_pdf_file)

print(f"PDF file '{output_pdf_file}' has been created.")