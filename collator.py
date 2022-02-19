
import fitz
import sys
import glob
import os
import shutil
import statistics


class MarkerComment:
    author: str
    type: str
    text: str
    rect: list[float]
    flags: int
    page: int
    original_raw_info: dict[str, str]


def main():
    # read cmd args
    try:
        pdf_dir = sys.argv[1]
        pdf_base = sys.argv[2]
        pdf_out = os.path.join(pdf_dir, "output.pdf")
    except:
        print("Two command line arguements required:\npython collator.py [directory] [base pdf]")

    # generate list of given marker pdfs
    pdf_list = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    pdf_list.remove(os.path.join(pdf_dir, pdf_base))
    if pdf_out in pdf_list:
        pdf_list.remove(pdf_out)

    print("Collating annotations on the following documents:")
    for pdf in pdf_list:
        print("  - " + pdf)

    # container for all annotations
    marking_annotations: list[MarkerComment] = []
    marking_marks: dict[str, list[float]] = {}

    # read each pdf and extract marks and comments
    for pdf_path in pdf_list:
        document = fitz.open(pdf_path)
        for page in document:
            for annotation in page.annots():
                if annotation.info["content"].strip().startswith("!#"):
                    marking_info = annotation.info["content"].strip().split()
                    question_number = marking_info[1]
                    question_mark = float("{:.2f}".format(float(marking_info[2])))
                    if question_number in marking_marks:
                        marking_marks[question_number].append(question_mark)
                    else:
                        marking_marks[question_number] = [question_mark]
                else:
                    comment = MarkerComment()
                    comment.author = annotation.info["title"].strip()
                    comment.text = annotation.info["content"].strip()
                    comment.type = annotation.type[1]
                    comment.rect = [annotation.rect[0], annotation.rect[1], annotation.rect[2], annotation.rect[3]]
                    comment.page = page.number
                    comment.flags = annotation.flags
                    comment.original_raw_info = annotation.info
                    marking_annotations.append(comment)
        document.close()

    # replace marker names
    origial_author_names = []
    for ma in marking_annotations:
        if ma.author not in origial_author_names:
            origial_author_names.append(ma.author)
    replacement_author_names = {}
    for i in range(len(origial_author_names)):
        replacement_author_names[origial_author_names[i]] = ("Marker #" + str(i+1))
    for ma in marking_annotations:
        ma.author = replacement_author_names[ma.author]

    print("A total of " + str(len(marking_annotations)) + " annotations have been extracted. There were " + str(len(origial_author_names)) + " authors which will be renamed.")
    print("An annotated copy is being generated based on:\n  - " + os.path.join(pdf_dir, pdf_base))

    # create new output pdf from template one
    shutil.copyfile(os.path.join(pdf_dir, pdf_base), pdf_out)

    # write comments to the output pdf
    document = fitz.open(pdf_out)
    for page in document:
        for ma in marking_annotations:
            if ma.page == page.number:
                if ma.type == "Text" or ma.type == "FreeText":
                    an = page.add_text_annot([ma.rect[0], ma.rect[1]], ma.text)
                elif ma.type == "Highlight":
                    an = page.add_highlight_annot(ma.rect)
                elif ma.type == "StrikeOut":
                    an = page.add_strikeout_annot(ma.rect)
                else:
                    print("Warning: unhandled annotation was ignored. Page: " + str(ma.page))
                    continue
                an.set_info(content=ma.text, title=ma.author, creationDate=ma.original_raw_info["creationDate"], modDate=ma.original_raw_info["modDate"])
                an.set_flags(ma.flags)
                an.update()

    # average the marks
    marking_marks_mean = {}
    marking_marks_total = 0
    for question_number in marking_marks:
        marking_marks_mean[question_number] = statistics.mean(marking_marks[question_number])
        marking_marks_total += marking_marks_mean[question_number]

    # write the total mark
    first_page = document[0]
    qa = first_page.add_text_annot([30.0, 30.0], "Overall mark: " + "{:.2f}".format(marking_marks_total))
    qa.set_colors({"stroke": (1.0, 0.0, 0.0), "fill": None})
    qa.set_info(title="Markers")
    qa.update()

    # write the individual marks
    for index, question_number in enumerate(marking_marks_mean):
        qa = first_page.add_text_annot([30.0, 80.0 + 20.0*index], "Question " + question_number + ": " + "{:.2f}".format(marking_marks_mean[question_number]))
        qa.set_info(title="Markers")
        qa.set_colors({"stroke": (1.0, 0.0, 0.0), "fill": None})
        qa.update()

    # save the annotated pdf
    document.saveIncr()
    document.close()

    print("Annotated pdf succesfully created:\n  - " + pdf_out)


if __name__ == "__main__":
    main()
