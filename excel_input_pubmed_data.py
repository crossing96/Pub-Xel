import xlwings as xw
import requests
from bs4 import BeautifulSoup
import re
import os

def input_pubmed_data():

    def remove_after_comma(s):
        parts = s.split(',', 1)  # Split the string at the first comma
        return parts[0]  # Return the part before the comma

    def remove_after_space(s):
        parts = s.split(' ', 1)  # Split the string at the first space
        if len(parts) > 1:
            return parts[0]  # Return the part before the space
        else:
            return s  # Return the whole string if there are no spaces

    #impact factor
    #impact factor 파일 업데이트 하면 맨앞 맨끝 문자 정리좀 해주셈

    # Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Build the full path to the file
    file_path = os.path.join(script_dir, 'impactfactor2022.txt')
    impactfactor2022 = open(file_path, "r", encoding="utf8")
    
    impactfactor2022 = str(impactfactor2022.read())
    impactfactor2022 = impactfactor2022.upper()
    impactfactor2022 = impactfactor2022.replace(": [","], ")
    impactfactor2022 = impactfactor2022.replace("'", "")
    impactfactor2022 = impactfactor2022.split("], ")
    impactfactor2022 = {impactfactor2022[i]: impactfactor2022[i + 1] for i in range(0, len(impactfactor2022), 2)}
    impactfactor2022.get("")

    #settings
    #header {column name : requested varaible}
    header = {"ref":"PMID","doi":"doi",
    "firstAuthor":"fa","author":"au","year":"yr","authorYear":"fayr",
    "journal":"jo","IF2022":"IF2022","title":"ti","abstract":"ab",
    "output1":"ou1","output2":"ou2","authors":"au2"}

    header2 = {}
    unidentifiedPMID=[]
    app = xw.apps.active
    wb = xw.books.active
    ws = xw.sheets.active


    # Extend the selection to the entire row
    rng = wb.app.selection

    ###identify relevant first row selection, and expand rng
    # Step 1: Identify column range of the selection
    column_range = (rng.column, rng.column + rng.columns.count - 1)
    # Step 2: Select a range of the first row of the sheet, ranging from the column range of the selection
    new_range = ws.range((1, column_range[0]), (1, column_range[1]))
    # Step 3: Extend the range to include the entire row
    entire_row_range = new_range.expand('right')
    # Step 4: Return the column range of the new range
    new_column_range = (entire_row_range.column, entire_row_range.column + entire_row_range.columns.count - 1)
    # Step 5: Extend the initial selection to include the columns identified in new_column_range, then select the new range
    extended_selection = ws.range((rng.row, new_column_range[0]), (rng.row + rng.rows.count - 1, new_column_range[1]))
    extended_selection.select()
    rng = wb.app.selection

    #identify headers
    for i in list(rng[0,:]):
        if ws[0,i.column-1].value in header:
            header2[header.get(ws[0,i.column-1].value)]=i.column-1

    rng_col_range = range(rng[:,0].row-1,rng[:,0].row-1+len(rng[:,0]))
    requested_ref_count = sum(1 for number in rng_col_range if number != 0) # length of rng_col_range but minus 1 if contain 0. 

    # # Disable screen updating
    # app.api.ScreenUpdating = False

    #input values
    for i in rng_col_range:
        if(i==0):
            continue
        PMID = ws[i,header2.get("PMID")].value

        a=PMID+1 # PMID가 숫자가 아닐 시 오류 발생시킴
        PMIDstring = str(PMID).rstrip('0').rstrip('.') # PMID to string
        
        url = "https://pubmed.ncbi.nlm.nih.gov/"+ PMIDstring +"/"
        response = requests.get(url)
        
        if(response.status_code!=200):
            unidentifiedPMID.append(PMIDstring)
            continue
        
        html_doc = response.text
        soup = BeautifulSoup(html_doc, 'html.parser')

        try:
            title = soup.find('meta', attrs={'name': 'citation_title'})['content']
        except Exception:
            title = 'NA'

        try:
            abstract_div = soup.find('div', {'id': 'abstract'}) # Find the div with id 'abstract'
            paragraphs = abstract_div.find_all('p') # Find all paragraph tags within the abstract div
            abstract = ' '.join(p.get_text() for p in paragraphs) # Extract and join the text from each paragraph
            abstract = abstract.replace('\n', ' ')
            abstract = re.sub(' +', ' ', abstract)
            abstract = abstract.lstrip()
            abstract = abstract.rstrip()
        except Exception:
            abstract = 'NA'

        try:
            doi=soup.find('meta', attrs={'name': 'citation_doi'})['content']
        except Exception:
            doi = 'NA'

        try:
            journal=soup.find('meta', attrs={'name': 'citation_publisher'})['content']
        except Exception:
            journal = 'NA'
        
        try:    
            # Find the div with class 'inline-authors'
            inline_authors_div = soup.find('div', {'class': 'inline-authors'})
            # Within 'inline_authors_div', find the div with class 'authors'
            authors_div = inline_authors_div.find('div', {'class': 'authors'})
            # Within 'authors_div', find the div with class 'authors-list'
            authors_list_div = authors_div.find('div', {'class': 'authors-list'})
            # Find all span tags within the authors_list_div
            span_elements = authors_list_div.find_all('span')
            # Initialize an empty string to hold the author names
            authors_text = ''
            # Loop through each span element

            for span in span_elements:
                if 'authors-list-item' in span['class']:
                    try:
                        # If the span has class 'authors-list-item', add the author name to the string
                        authors_text += span.find('a', {'class': 'full-name'}).get_text(strip=True)
                    except Exception:
                        try:
                            authors_text += span.find('span', {'class': 'full-name'}).text # for PMID 34967848
                        except Exception:
                            pass

                elif 'comma' in span['class']:
                    # If the span has class 'comma', add a comma to the string
                    authors_text += ', '
                elif 'semicolon' in span['class']:
                    # If the span has class 'semicolon', add a semicolon to the string
                    authors_text += '; '
                    
            authors=authors_text

        except Exception:
            authors = 'NA'

        try:
            authortemp=soup.find('meta', attrs={'name': 'citation_authors'})['content']
            authortemp=authortemp.split(";")[0]
            if authortemp != " ":
                firstauthor = authortemp
            else:
                firstauthor = remove_after_comma(authors.replace('; ', ', '))
        except Exception:
            authortemp = 'NA'

        try:
            date = soup.find('meta', attrs={'name': 'citation_date'})['content']
        except Exception:
            date = 'NA'
        
        if(date.count('/') == 2):
            year=date.rpartition('/')[-1]
        else:
            year=remove_after_space(date)

        firstauthoryear = str(firstauthor) + ", " + str(year)

        try:
            iftemp = impactfactor2022.get(journal.upper())
            IF2022=impactfactor2022.get(journal.upper())[0:iftemp.find(",")]
        except Exception:
            IF2022='NA'
        
        if header2.get("doi",-1)>=0:
            ws[i,header2.get("doi")].value = doi
        if header2.get("au2",-1)>=0:
            ws[i,header2.get("au2")].value = authors
        if header2.get("au",-1)>=0:
            ws[i,header2.get("au")].value = authors
        if header2.get("fa",-1)>=0:
            ws[i,header2.get("fa")].value = firstauthor
        if header2.get("ti",-1)>=0:
            ws[i,header2.get("ti")].value = title
        if header2.get("ab",-1)>=0:
            ws[i,header2.get("ab")].value = abstract
        if header2.get("jo",-1)>=0:
            ws[i,header2.get("jo")].value = journal
        if header2.get("yr",-1)>=0:
            ws[i,header2.get("yr")].value = year
        if header2.get("fayr",-1)>=0:
            ws[i,header2.get("fayr")].value = firstauthoryear
        if header2.get("IF2022",-1)>=0:
            ws[i,header2.get("IF2022")].value = IF2022
        if header2.get("ou1",-1)>=0:
            ws[i,header2.get("ou1")].value = firstauthor + ", " + year + ". PMID: " + PMIDstring + ", " + journal + " (IF: " + IF2022 + "). " + title

    # # Enable screen updating
    # app.api.ScreenUpdating = True

    if True: 
        print("Number of requested references: "+str(requested_ref_count))
        print("Unidentified references: " + str(unidentifiedPMID))
