require 'docx'
require 'nokogiri'

file = './document.docx'
doc = Docx::Document.open(file)

html_doc = Nokogiri::HTML::DocumentFragment.parse("")

# Wordの各段落をHTMLに変換
doc.paragraphs.each do |p|
  html_doc.add_child("<p>#{p}</p>")
end

File.open("output.html", "w") { |file| file.write(html_doc.to_html) }
