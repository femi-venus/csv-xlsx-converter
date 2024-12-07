
require 'caxlsx'
require 'csv'

csv_file_path = '/home/femi/training/xlsx-converter/xlsx-csv.csv'
image_file_path = '/home/femi/training/xlsx-converter/car-images/'

Axlsx::Package.new do |p|
  p.workbook.add_worksheet(name: "Car Data") do |sheet|

    CSV.foreach(csv_file_path, headers: true) do |row|

        if sheet.rows.empty?
            sheet.add_row row.headers ,b:true
         end
          id = row['id']
      sheet.add_row row.fields,height:100
      sheet.add_image(image_src: image_file_path + "#{id}.jpeg", noSelect: true, noMove: true) do |image|
        image.width = 100
        image.height = 100
        image.start_at 7, ("#{id}").to_i

      end
    end
  end

  p.serialize('csv-xlsx.xlsx')
end

puts "Excel file created at 'csv-xlsx.xlsx'!"

#///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# require 'caxlsx'
# require 'csv'

# csv_file_path = '/home/femi/training/xlsx-converter/xlsx-csv.csv'

# def cars_data_display(csv_file_path)
# header_written = false
# Axlsx::Package.new do |p|
#   p.workbook.add_worksheet(name: "Car Data") do |sheet|

#     CSV.foreach(csv_file_path, headers: true) do |row|

#        unless header_written
#             sheet.add_row row.headers ,b:true
#             header_written = true
#          end
#           id = row['id']
#           image = row['image']
#       sheet.add_row row.fields[0..-2],height:100
#       sheet.add_image(image_src: image , noSelect: true, noMove: true) do |image|
#         image.width = 100
#         image.height = 100
#         image.start_at 7, ("#{id}").to_i

#       end
#     end
#   end
#   p.serialize('csv-xlsx.xlsx')
# end
# end

# cars_data_display(csv_file_path);
# puts "Excel file created at 'csv-xlsx.xlsx'!"




#/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# require "nokogiri"
# require "csv"

# xml_coordinates_content = <<-XML
#  <?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId3"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId4"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId5"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId6"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId7"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId8"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId10"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId11"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="" descr=""></xdr:cNvPr><xdr:cNvPicPr><a:picLocks noSelect="1" noChangeAspect="1" noMove="1" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r ="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId12"></a:blip><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>
# XML

# coord_doc = Nokogiri::XML(xml_coordinates_content)

# coordinates = {}

# coord_doc.xpath("//xdr:oneCellAnchor", "xdr" => "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing").each do |anchor|
#   col = anchor.at_xpath("xdr:from/xdr:col", "xdr" => "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing").text
#   row = anchor.at_xpath("xdr:from/xdr:row", "xdr" => "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing").text

#   r_id = anchor.at_xpath(".//a:blip/@r:embed", "a" => "http://schemas.openxmlformats.org/drawingml/2006/main", "r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships").to_s

#   coordinates[r_id] = { col: col.to_i, row: row.to_i } unless r_id.empty?
# end

# # puts coordinates

# xml_images_content = <<-XML
# <?xml version="1.0" encoding="UTF-8"?>
# <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
# <Relationship Target="../media/image1.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId3"/>
# <Relationship Target="../media/image2.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId4"/>
# <Relationship Target="../media/image3.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId5"/>
# <Relationship Target="../media/image4.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId6"/>
# <Relationship Target="../media/image5.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId7"/>
# <Relationship Target="../media/image6.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId8"/>
# <Relationship Target="../media/image7.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId9"/>
# <Relationship Target="../media/image8.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId10"/>
# <Relationship Target="../media/image9.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId11"/>
# <Relationship Target="../media/image10.jpeg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="rId12"/>
# </Relationships>
# XML
# img_doc = Nokogiri::XML(xml_images_content)

# namespace = { "ns" => "http://schemas.openxmlformats.org/package/2006/relationships" }

# images = {}

# img_doc.xpath("//ns:Relationship", namespace).each do |relationship|
#   id = relationship["Id"]
#   target = relationship["Target"]
#   images[id] = target if id && target
# end

# # puts images

# merged_output = []

# coordinates.each do |r_id, coord|
#   if images.key?(r_id)
#     merged_output << {
#       col: coord[:col],
#       row: coord[:row],
#       image: images[r_id],
#     }
#   end
# end

# # puts merged_output

# cardata = []
# CSV.foreach("xlsx-csv.csv", headers: true, header_converters: :symbol) do |row|
#   headers ||= row.headers
#   cardata << row.to_h
# end
# #  puts cardata

# new_data = cardata.each do |data|
#   merged_output.each do |op|
#     if op[:row] == data[:id].to_i
#       data[:image] = op[:image]
#       break
#     end
#   end
# end

# puts new_data.inspect


# require 'rubyXL'

