require "roo"
require "csv"
require "zip"
require "fileutils"

xlsx_file_path = "/home/femi/training/xlsx-converter/csv-xlsx.xlsx"

def convert_xlsx_to_csv_and_extract_images(xlsx_file_path)
  output_csv_path = "xlsx-csv.csv"
  output_image_folder = "/home/femi/training/xlsx-converter/output_images"

  xls_file = Roo::Excelx.new(xlsx_file_path)
  images = xls_file.image_files
  CSV.open(output_csv_path, "w") do |csv|
    (1..xls_file.last_row).each do |i|
      csv << xls_file.row(i)
    end
  end
  puts "CSV file created at '#{output_csv_path}'!"

  FileUtils.mkdir_p(output_image_folder)

  Zip::File.open(xlsx_file_path) do |zip|
    zip.each do |entry|
      if entry.name.start_with?("xl/media/") && entry.name.end_with?(".jpeg")
        output_file_path = File.join(output_image_folder, File.basename(entry.name))
        File.open(output_file_path, "wb") do |file|
          file.write(entry.get_input_stream.read)
        end
      end
    end
  end
  puts "Images extracted to '#{output_image_folder}'!"
end

convert_xlsx_to_csv_and_extract_images(xlsx_file_path)