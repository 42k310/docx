require 'zip'
require 'nokogiri'

file = './document.docx'
body_xml_content = nil
comments_xml_content = nil

# ZIPアーカイブ内のコメントXMLファイルを読み込む
Zip::File.open(file) do |zip_file|
  zip_file.each do |entry|
    if entry.name == 'word/comments.xml'
      comments_xml_content = entry.get_input_stream.read
    # elsif entry.name == 'word/document.xml'
    #   body_xml_content = entry.get_input_stream.read
    end
  end
end

# document.bodyを取って解析することも可能
# if body_xml_content
#   doc = Nokogiri::XML(body_xml_content)
#   doc.xpath('//w:t').each do |text_element|
#     puts text_element.text
#   end
# else
#   puts "No document content found in the document."
# end

if comments_xml_content
  comments_doc = Nokogiri::XML(comments_xml_content)
  comments_doc.xpath('//w:comment').each do |comment_node|
    author = comment_node['w:author']
    date = comment_node['w:date']
    text = comment_node.text.strip

    puts "Author: #{author}, Date: #{date}, Comment: #{text}"
  end
else
  puts "No comments found in the document."
end
