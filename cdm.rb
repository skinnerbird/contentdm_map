require 'marc'
require 'marc_tools'
require 'spreadsheet'

# SCL Metadata Application Profile processor
# Generates .xls in accordance with specification
# for import into Contentdm
# Original: 09/19/2011, Updated: 06/11/2012
# Author: Mark Cooper, Modified gy: Geoffrey Skinner

=begin
SET CRITERIA / CONSTRAINTS FOR FIELDS
=end

module MARC
	class DataField
		# monekey patch: space delimit subfields
		def value
  			return(@subfields.map {|s| s.value}.join ' ')
		end
	end
end

input_file = 'CDM_MARC_conv.mrc'
NUM_HEADERS = 92
spec = {}
File.open('spec.txt').each_line do |l|
	header, width = l.split('|')
	spec[header] = width.strip
end

raise "Incorrect number of headers" unless spec.size == NUM_HEADERS

num_records     = 0
cdm             = []

keywords = {
	:source_of_title    	=> 'title (from|supplied)',
	:publication_note   	=> '(postcard|map|advertis.*|publish.*)',
	:source_of_alt_title	=> '(other title|title on)',
	:reproduction       	=> 'reproduced',
	:creator_note       	=> 'photographed',
	:condition				=> '(torn|missing)',
	:other_notes			=> '(signed|written|printed|letters|numbers|signed|written|printed)',
	:numbering				=> '(letters|numbers)',
	:signatures				=> '(signed|written|printed)',
	:source_note			=> 'originally',
	:marks_stamps			=> '(mark|stamped)',
	:digitization			=> 'scanned',
	:ownership				=> '[Oo]wned by [Ss]onoma [Cc]ounty [Ll]ibrary',
}
note_string = '(' + keywords.values.join('|') + ')'

=begin
ASSEMBLE RECORD DATA INTO ARRAY (CDM) OF HASHES (CDM_DATA ELEMENTS)
=end

MARC::ForgivingReader.new(input_file).each do |r|
	num_records += 1
	record = MarcTools::MarcRecordWrapper.new(r)
	cdm_data = Hash.new
	begin
		cdm_data['Title'] = MarcTools::MarcFieldWrapper.new(record['245']).subfield_string(' ', 'c', 'h').strip
		cdm_data['Statement of Responsibility'] = record['245']['c'] # REC MUST HAVE 245
		cdm_data['Creator'] = MarcTools::MarcFieldWrapper.new(record['100']).subfield_string(' ', 'e').strip rescue nil
		cdm_data['Other Titles'] = record.grab('246').map(&:value).join(';').strip
		cdm_data['Contributor (Person)'] = record.grab('700').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', 'e', 't').strip}.join(';').strip
		cdm_data['Contributor (Corporate)'] = record.grab('710').map(&:value).join(';').strip
		cdm_data['Contributor (Conference or Meeting)'] = record.grab('711').map(&:value).join(';').strip
		cdm_data['Description'] = record.grab('520').map(&:value).join(';').strip
		cdm_data['Biographical or Historical Note'] = record['545'].value rescue nil
		cdm_data['Date Created or Published'] = record['260']['c'] rescue nil # date1 ???
		cdm_data['Date Reprinted'] = record.dtst == 'r' ? record.date1 : ''
		cdm_data['Date Created or Published (clean -- 1)'] = record.dtst == 'r' ? record.date1 : ''
		cdm_data['Date Created or Published (clean -- 2)'] = record.dtst == 'r' ? record.date2 : ''
		cdm_data['Publisher location (Original)'] = record['260']['a'].gsub(/^\[/, '').gsub(/\s*\W$/, '') rescue nil
		cdm_data['Publisher (Original)'] = record['260']['b'] rescue nil
		cdm_data['Table Of Contents '] = record.grab('505').map(&:value).join(';').strip
		cdm_data['Item Type (Original)'] = record.type
		cdm_data['Item Physical Format or Genre'] = record.grab('655').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip	
		cdm_data['Number of Parts or Pages'] = record['300']['a'].gsub(/:\s*$/, '').strip rescue nil
		cdm_data['Other Physical Details'] = record['300']['b'].gsub(/;\s*$/, '').strip rescue nil
		cdm_data['Dimensions'] = record['300']['c'] rescue nil
		cdm_data['Language'] = record.grab('041').map(&:subfields).join(';').strip # TEST
		cdm_data['Sonoma Heritage Collections Theme'] = 'INPUT' # TEMPLATE
		cdm_data['Subject (Person)']    = record.find_all{|f| f.tag == '600' and f.indicator1 =~ '/(0|1)/'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Family)']    = record.find_all{|f| f.tag == '600' and f.indicator1 == '3'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Corporate Body)'] = record.grab_by_value('610').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Meeting or Event)']   = record.grab_by_value('611').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Title)']   = record.grab_by_value('630').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Time Period)']       = record.grab_by_value('648', 'fast$').map{ |f|  MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Topical - FAST)']    = record.grab_by_value('650').map{ |f|  MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['Subject (Geographic Feature)']   = record.grab_by_value('651').map{ |f|  MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ', '2').gsub(/ -- $/, '')}.join(';').strip
		cdm_data['City or Town'] = '' # TEMPLATE
		cdm_data['District'] = '' # TEMPLATE
		cdm_data['California County'] = '' # TEMPLATE
		cdm_data['State or Province'] = '' # TEMPLATE
		cdm_data['Country'] = '' # TEMPLATE
		cdm_data['Geographic Metadata Source'] = '' # TEMPLATE
		cdm_data['Street Address'] = record['939']['a'] rescue nil
		cdm_data['Geolocation'] = record['598']['a'] rescue nil
		cdm_data['Geocoding Status'] = record['598'].value rescue nil
		cdm_data['Map Zoom Level'] = '' # TEMPLATE
		cdm_data['Creator Note'] = record.grab_by_value('500', keywords[:creator_note]).map(&:value).join(';').strip
		cdm_data['Source of Title'] = record.grab_by_value('500', keywords[:source_of_title]).map(&:value).join(';').strip
		cdm_data['Source of Title Variations'] = record.grab_by_value('500', keywords[:source_of_alt_title]).map(&:value).join(';').strip
		cdm_data['Original Creation or Publication Details'] = record.grab_by_value('500', keywords[:publication_note]).map(&:value).join(';').strip
		cdm_data['Reproduction Details (Physical Items)'] = record.grab_by_value('500', keywords[:reproduction]).map(&:value).join(';').strip
		cdm_data['Other Notes'] = record.grab_by_value('500', keywords[:other_notes]).map(&:value).join(';').strip
		cdm_data['Unidentified notes'] = record.find_all {|f| f.tag == '500' and f.value !~ /#{note_string}/i}.map(&:value).join(';').strip
		cdm_data['Source Note'] = record.grab_by_value('500', keywords[:source_note]).map(&:value).join(';').strip
		cdm_data['Physical Condition'] = record.grab_by_value('500', keywords[:condition]).map(&:value).join(';').strip
		cdm_data['Referenced By'] = record.find_all {|f| f.tag == '700' and f['t']}.map(&:value).join(';').strip
		cdm_data['Related Periodical, Etc.'] = record['730'].value rescue nil
		cdm_data['Related Publication (Book)'] = record['787'].value rescue nil
		cdm_data['Related Publication (Web)''] = record.find_all {|f| f.tag == '856'  and f.indicator1 =~ /(2)/' and f['u'] !~ /\.jpg/}.map{|f| f['u']}.join(';').strip
		cdm_data['Map Scale'] = record['034'].value rescue nil
		cdm_data['Collection Name'] = record.grab('8[03]0').map{ |f| f.value.gsub(/;/, ':') }.join(';').strip
		cdm_data['URI'] = record.find_all {|f| f.tag == '856' and f.indicator1 =~ '/(0|1)/' and f['u'] !~ /\.jpg/}.map{|f| f['u']}.join(';').strip
		cdm_data['Collection Guide'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Full Fext'] = '' # TEMPLATE
		cdm_data['CONTENTdm Collection Name'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Contributing Organization'] = record.grab_by_value('500', keywords[:ownership]).map(&:value).join(';').strip
		cdm_data['Rights Management'] = record['540']['a'] rescue nil
		cdm_data['Copyright Date'] = record.copyright_date rescue nil
		cdm_data['Project Affiliation'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Fiscal Sponsor'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Local System Identifier'] = record['996'].value rescue nil
		cdm_data['SCH Identifier'] = 'SUPPLIED' # TEMPLATE
#		cdm_data['Item Call Number'] = record.find_all {|f| f.tag == '856' and f['3']}.map{|f| f['3'].split(';')[0]}.join(';').strip
		cdm_data['Item Call Number'] = = record.grab('949').map{ |f| f['m'] }.join(';').strip
		cdm_data['Housing Loc of Physical Item(s)'] = record.grab('949').map{ |f| f['d'] }.join(';').strip
		cdm_data['Location of Originals'] = record['535'].value rescue nil
		cdm_data['Immediate Source of Acquisition'] = record['541'].value rescue nil
		cdm_data['Provenance'] = record['561'].value rescue nil
		cdm_data['Internal Note'] = record['590'].value rescue nil
		cdm_data['Digitization Note'] = record.grab_by_value('500', keywords[:digitization]).map(&:value).join(';').strip
		cdm_data['Digitizing Agency'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Date Digitized'] = '' # TEMPLATE
		cdm_data['Item Digital Format'] = '' # TEMPLATE
		cdm_data['Archival Disc File ID'] = record.find_all {|f| f.tag == '856' and f['u']}.map{|f| f['u'].scan(/[A-Za-z]?\d+.?[A-Za-z]?\.jpg/)}.join(';').strip
		cdm_data['Master File Data Quality'] = record['909'].value rescue nil
		cdm_data['Master File Size'] = '' # TEMPLATE
		cdm_data['Master File Format'] = '' # TEMPLATE
		cdm_data['Master File Bit Depth'] = '' # TEMPLATE
		cdm_data['Master File Resolution'] = '' # TEMPLATE
		cdm_data['Master File Compression'] = '' # TEMPLATE
		cdm_data['Master File Width'] = '' # TEMPLATE
		cdm_data['Master File Height'] = '' # TEMPLATE
		cdm_data['Master File Photometric Interpretation'] = '' # TEMPLATE
		cdm_data['Master File Software'] = '' # TEMPLATE
		cdm_data['Master File System'] = '' # TEMPLATE
		cdm_data['Master File Checksum'] = '' # TEMPLATE
		cdm_data['Object File Name'] = '' # TEMPLATE
		cdm_data['OCLC#'] = record['001'].value rescue nil

		cdm_data.each { |k, v| cdm_data[k] = ' ' if v.nil? or v.empty? }
		raise "Consistency error: keys are not equal." unless cdm_data.keys == spec.keys
		cdm << cdm_data
	rescue Exception => ex
		puts record['001'].value + ' -- ' + ex.backtrace.join("\n")
	end
end

puts "RECORDS READ: " + num_records.to_s
cdm.sort_by! { |cdm_data| cdm_data['Title'] }

=begin
CREATE THE SPREADSHEET
=end

format = {
	:weight           => :bold,
	:color            => :white,
	:pattern_fg_color => :blue, 
	:pattern          => 1,
	:horizontal_align => :center,
	:vertical_align   => :center,
}
book = Spreadsheet::Workbook.new
transform = book.create_worksheet
transform.name = 'transform'
transform.row(0).concat spec.keys
transform.row(0).default_format = Spreadsheet::Format.new(format)
(0..spec.keys.size - 1).each { |i| transform.column(i).width = spec.values[i] }
cdm.each_with_index { |h, idx| idx += 1; transform.row(idx).replace(h.values.map { |v| v.encode('UTF-8') }) }
book.write "cdm-#{Time.now.strftime('%Y%m%d')}.xls"

# worksheet.row(idx + 1).replace attributes.map{ |v| v.encode('UTF-8') }