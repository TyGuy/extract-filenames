# need to run:
# gem install yomu
# before running this

require 'yomu'

def new_filename(original_filename)
  yomu = Yomu.new(original_filename)
  text = yomu.text
  extension = original_filename.split('.').last

  new_file = text[0..250].gsub(/\n/, ' ').gsub(/\s+/, ' ').tr('/', '-').gsub('Sheet1', '').strip[0..50]
  new_file = "#{new_file}.#{extension}"

  puts new_file
  new_file
end

def rename_file(original_filename)

  puts "-----------------------------------------"
  puts "Attempting to rename #{original_filename}"

  new_file = new_filename(original_filename)
  puts new_file

  if new_file
    file_exists = File.exist?(new_file)
    # puts file_exists

    if file_exists
      # puts "file exists!"
      parts = new_file.split('.')
      base = parts.first
      extension = parts.last
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
    filenames = Dir.glob('*.xls')
  elsif (original_filename = ARGV[0])
    filenames = [original_filename]
  else
    puts "usage: ruby read_word_doc.rb <filename>"
    return
  end

  filenames.each do |orig_file|
    rename_file(orig_file)
  end

  puts 'DONE.'
end


main if __FILE__ == $PROGRAM_NAME
