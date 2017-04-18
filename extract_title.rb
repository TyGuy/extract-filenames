require 'rexml/document'
include REXML

def new_filename(original_filename)
  unzipped_dir = original_filename.split('.').first + '_unzipped'

  zipfile = "#{original_filename}.zip"
  `cp "#{original_filename}" "#{zipfile}"`

  `unzip "#{zipfile}" -d "#{unzipped_dir}"`

  new_filename = try_find_filename(unzipped_dir)

  if new_filename
    extension = original_filename.split('.').last
    new_filename = "#{new_filename}.#{extension}"
    # puts "new filename: #{new_filename}"
  else
    puts 'could not find filename at all'
  end

  `rm -r "#{unzipped_dir}"`
  `rm "#{zipfile}"`

  new_filename
end

def try_find_filename(unzipped_dir)
  data_file_path = "#{unzipped_dir}/docProps/core.xml"
  name_from_data_file = search_for_title(data_file_path)

  if name_from_data_file && !name_from_data_file.empty?
    puts "FROM_METADATA"
    return name_from_data_file
  end

  document_file_path = "#{unzipped_dir}/word/document.xml"
  name_from_document = search_for_document(document_file_path)

  return unless name_from_document

  puts "FROM_DOCUMENT"
  name_from_document
end

def search_for_document(document_file_path)
  xmlfile = File.new(document_file_path)
  xmldoc = Document.new(xmlfile)

  first_text = nil

  xmldoc.elements.each('//w:t') do |text|
    first_text = text.text
    break
  end

  first_text && first_text[0..30]
end

def search_for_title(data_file_path)
  content = ''

  File.open(data_file_path, 'r') do |file|
    file.each_line do |line|
      content << line
    end
  end

  title = content.match(/\<dc\:title\>(.*)\<\/dc\:title\>/)[1]
  puts title

  title
end

def rename_file(original_filename)

  puts "-----------------------------------------"
  puts "Attempting to rename #{original_filename}"

  new_file = new_filename(original_filename)
  puts new_file

  if new_file
    file_exists = `[ ! -f "#{new_file}" ]`
    puts file_exists

    if file_exists
      puts "file exists!"
      base, extension = new_file.split('.')
      base = "#{base}_#{Time.now.to_f}"
      new_file = [base, extension].join('.')
    end

    puts "Renaming #{original_filename} to #{new_file}"
    `mv "#{original_filename}" "#{new_file}"`

  else
    puts "could not rename #{original_filename}"
  end
  puts "-----------------------------------------"
end

def main
  if ARGV.include?('--all')
    filenames = Dir.glob('*.docx')
  elsif (original_filename = ARGV[0])
    filenames = [original_filename]
  else
    puts "usage: ruby extract_title.rb <filename>"
    return
  end

  filenames.each do |orig_file|
    rename_file(orig_file)
  end

  puts 'DONE.'
end


main if __FILE__ == $PROGRAM_NAME
