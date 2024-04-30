import PyPDF2
import re

pdf_file = r"C:\Users\cbyington\Documents\HummingBirdPics\Binder1.pdf"
output_file = r"C:\Users\cbyington\Documents\HummingBirdPics\matter_numbers.txt"

def extract_content_from_page(page_number):
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        if 0 < page_number <= len(pdf_reader.pages):
            page = pdf_reader.pages[page_number - 1]
            content = page.extract_text()
            return strip_non_alphabetical(content)
        else:
            return "Invalid page number."

def extract_content_from_range(start_page, end_page):
    with open(pdf_file, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        content = ""
        for page_number in range(start_page, end_page + 1):
            if 0 < page_number <= len(pdf_reader.pages):
                page = pdf_reader.pages[page_number - 1]
                page_content = page.extract_text()
                content += strip_non_alphabetical(page_content) + ";"  # Strip non-alphabetical characters and add to content
            else:
                content += f"Invalid page number: {page_number}\n"
        return content

def get_pages_from_input(page_input):
    # Extract individual pages or ranges using regex
    pages = []
    page_ranges = re.findall(r"(\d+)\s*-\s*(\d+)", page_input)
    for start, end in page_ranges:
        pages.extend(range(int(start), int(end) + 1))
    individual_pages = re.findall(r"\d+", page_input)
    pages.extend(map(int, individual_pages))
    return sorted(set(pages))

def strip_non_alphabetical(text):
    # Remove all alphabetical and special characters from the text and trailing whitespace
    text = re.sub(r"[^0-9\s]+", "", text).strip()
    lines = text.splitlines()[:-4]
    first_numbers = [re.search(r"\d+", line).group() for line in lines if re.search(r"\d+", line)]
    result = ";".join(first_numbers)
    return result 

def count_matter_numbers(content):
    matter_numbers = content.split(';')
    return len([number for number in matter_numbers if number.strip()])

if __name__ == "__main__":
    page_input = input("Enter the page number(s) or range (press 0 to continue if you didn't finish last time. Entering new pages will overwrite the numbers you had left!): ")
    if page_input == "0":
        exit()

    # Check if the first character is a dash
    if page_input.startswith("-"):
        print("Invalid input. Page numbers cannot be negative. Please try again.")
        exit()

    pages = get_pages_from_input(page_input)

    if pages:
        if len(pages) == 1:
            page_number = pages[0]
            content = extract_content_from_page(page_number)

            # Write the content to the output file
            with open(output_file, 'w') as file:
                file.write(content + "\n")
        else:
            start_page = pages[0]
            end_page = pages[-1]
            content = extract_content_from_range(start_page, end_page)

            # Write the content to the output file with a maximum of 20 matter numbers per line
            matter_numbers = content.split(';')
            with open(output_file, 'w') as file:
                lines = []
                current_line = []
                for i, matter_number in enumerate(matter_numbers, start=1):
                    current_line.append(matter_number)
                    if i % 20 == 0:
                        lines.append(';'.join(current_line))
                        current_line = []
                if current_line:
                    lines.append(';'.join(current_line))
                file.write('\n'.join(lines) + "\n")
    else:
        print("Invalid input. Please enter valid page number(s) or range.")
