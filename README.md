# PowerPoint_Markdown_converter
This is a CLI application that converts Powerpoint slides into Markdown. I built this with the idea of importing the final output to Roam Research. This influenced a lot of the design decisions (and probably will in future versions.) 

## Usage 
to use the app simply type this command in your terminal:
`python pptx2markdown.py [powerpoint filename]`

A folder named '[ppt filename]_converted' will be generated. in it, you'll find the generated markdown file and images folder.

## Output
The format outpul looks like this:

- Slide number 1
	- content
- Slide number 2
	- content
- Slide number
	- content

It doesn't matter how the powerpoint is structured, this is the final output. Whether there are bullet points and indentation or not, it doesn't matter. More formatting can be applied in Roam itself. 

If there are images in the Powerpoint, they will be saved in a specific folder with the number of the slide they belong to in the title. 

Feedback is appreciated! 