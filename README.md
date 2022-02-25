# Slidepack Generator
These functions support the creation of slidepacks for common analytical use-cases. It builds a wrapper around the python-pptx library and provides a very simple and limited way-in, with a focus on headers, text and images - which is anticipated to cover most Data uses cases where chart images and accompanying text are required.

## Suggested workflow
pres = Presentation('filename.pptx')
 _slide2_idx = add_slide(...)
 _slide3_idx = add_slide(...)
 ...
 pres.save('output_file.pptx')

(you can use the slide idx to generate a contents page).
