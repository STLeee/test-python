# from docx2pdf import convert
import pypandoc

from test_docx import main as docx_main


def main():
    # Call the main function from test_docx to create the document
    docx_main()

    # convert("demo.docx", "demo.pdf")
    pypandoc.convert_file(
        "demo.docx",
        "pdf",
        outputfile="demo.pdf",
        extra_args=[
            "--pdf-engine=xelatex",
            "--variable=mainfont:Microsoft JhengHei",
        ],
    )


if __name__ == "__main__":
    main()
    print("Document created successfully.")
