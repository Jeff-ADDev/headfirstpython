file_name = "./chapter_4/textfiles/example.txt"
content = "This is some text that we want to save in the file."

# Open the file with 'write' mode
with open(file_name, 'w') as file:
    # Write the content to the file
    file.write(content)