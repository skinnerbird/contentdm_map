require 'marc'
require 'marc_tools'
require 'spreadsheet'

# SCL Metadata Application Profile processor
# Generates .xls in accordance with specification
# for import into Contentdm
# Original: 09/19/2011, Updated: 11/28/2011
# Author: Mark Cooper

=begin
SET CRITERIA / CONSTRAINTS FOR FIELDS
=end

NUM_HEADERS = 72
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

MARC::ForgivingReader.new('PhotoRecordConversions_CDM_ready_TEST.mrc').each do |r|
	num_records += 1
	record = MarcTools::MarcRecordWrapper.new(r)
	cdm_data = Hash.new
	begin
		cdm_data['Creator'] = MarcTools::MarcFieldWrapper.new(record['100']).subfield_string(' ', 'e').strip rescue nil
		cdm_data['Creator Role'] = record['100']['e'] rescue nil
		cdm_data['Title'] = MarcTools::MarcFieldWrapper.new(record['245']).subfield_string(' ', 'c', 'h').strip
		cdm_data['Statement of Responsibility'] = record['245']['c'] # REC MUST HAVE 245
		cdm_data['Alternative Title'] = record.grab('246').map(&:value).join(';').strip
		cdm_data['Place of Publication (Original)'] = record['260']['a'].gsub(/\s*\W$/, '') rescue nil
		cdm_data['Publisher (Original)'] = record['260']['b'] rescue nil
		cdm_data['Creation/Publication Date'] = record['260']['c'] rescue nil # date1 ???
		cdm_data['Copyright Date'] = record.copyright_date
		cdm_data['Reprint Date'] = record.dtst == 'r' ? record.date1 : ''
		cdm_data['Creation/Publication Date (clean)'] = record.dtst == 'r' ? record.date1 : ''
		cdm_data['Creation/Publication Date (2) (clean)'] = record.dtst == 'r' ? record.date2 : ''
		cdm_data['Object Description'] = record.grab('520').map(&:value).join(';').strip
		cdm_data['Biographical/Historical Note'] = record['545'].value rescue nil
		cdm_data['Table Of Contents '] = record.grab('505').map(&:value).join(';').strip
		cdm_data['Creator Note'] = record.grab_by_value('500', keywords[:creator_note]).map(&:value).join(';').strip
		cdm_data['Number of Images/Parts'] = record['300']['a'].gsub(/:\s*$/, '').strip rescue nil
		cdm_data['Other Physical Details'] = record['300']['b'].gsub(/;\s*$/, '').strip rescue nil
		cdm_data['Dimensions'] = record['300']['c'] rescue nil
		cdm_data['Language'] = record.grab('041').map(&:subfields).join(';').strip # TEST
		cdm_data['Source of Title'] = record.grab_by_value('500', keywords[:source_of_title]).map(&:value).join(';').strip
		cdm_data['Source of Title Variations'] = record.grab_by_value('500', keywords[:source_of_alt_title]).map(&:value).join(';').strip
		cdm_data['Original Creation/Publication Details'] = record.grab_by_value('500', keywords[:publication_note]).map(&:value).join(';').strip
		cdm_data['Reproduction Details (physical item)'] = record.grab_by_value('500', keywords[:reproduction]).map(&:value).join(';').strip
		cdm_data['Location of Originals'] = record['535'].value rescue nil
		cdm_data['Collection Guide'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Owning Institution'] = record.grab_by_value('500', keywords[:ownership]).map(&:value).join(';').strip
		cdm_data['Copyright Status'] = record['540']['a'] rescue nil
		cdm_data['Reuse and Reproduction Restrictions'] = 'SUPPLIED' # TEMPLATE
		cdm_data['Immediate Source of Acquisition'] = record['541'].value rescue nil
		cdm_data['Provenance'] = record['561'].value rescue nil
		cdm_data['Physical Condition'] = record.grab_by_value('500', keywords[:condition]).map(&:value).join(';').strip
		cdm_data['Signatures and Inscriptions'] = record.grab_by_value('500', keywords[:signatures]).map(&:value).join(';').strip
		cdm_data['Numbers or Letters on Originals'] = record.grab_by_value('500', keywords[:numbering]).map(&:value).join(';').strip
		cdm_data['Markings and Stamps'] = record.grab_by_value('500', keywords[:marks_stamps]).map(&:value).join(';').strip
		cdm_data['Other Statements of Responsibility'] = 'N/A'
		cdm_data['Source Note'] = record.grab_by_value('500', keywords[:source_note]).map(&:value).join(';').strip
		cdm_data['Local Note'] = record['590'].value rescue nil
		cdm_data['Geocoding status'] = record['598'].value rescue nil
		cdm_data['Subject (Person - LCNAF)']    = record.find_all{|f| f.tag == '600' and f.indicator1 =~ /(0|1)/ and f.indicator2 != '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Person - Local)']    = record.find_all{|f| f.tag == '600' and f.indicator1 =~ /(0|1)/ and f.indicator2 == '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Family - LCNAF)']    = record.find_all{|f| f.tag == '600' and f.indicator1 == '3' and f.indicator2 != '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Family - Local)']    = record.find_all{|f| f.tag == '600' and f.indicator1 == '3' and f.indicator2 == '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Corporate - LCNAF)'] = record.find_all{|f| f.tag == '610' and f.indicator2 != '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Corporate - Local)'] = record.find_all{|f| f.tag == '610' and f.indicator2 == '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Meeting - LCNAF)']   = record.find_all{|f| f.tag == '611' and f.indicator2 != '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Meeting - Local)']   = record.find_all{|f| f.tag == '611' and f.indicator2 == '7'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Time Period)']       = record.grab_by_value('648', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Topical - FAST)']    = record.grab_by_value('650', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Topical - LOCAL)']   = record.grab_by_value('650', 'local$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Geographic FAST)']   = record.grab_by_value('651', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Geographic LCNAF)']  = record.grab_by_value('651', 'lcnaf$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Subject (Geographic local)']  = record.grab_by_value('651', 'local$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' -- ').strip}.join(';').strip
		cdm_data['Street Address'] = record['939']['a'] rescue nil
		cdm_data['Location Coordinates'] = record['938']['a'] rescue nil
		cdm_data['Genre/Form'] = record.grab('655').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip	
		cdm_data['Contributor (Person)'] = record.grab('700').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', 'e', 't').strip}.join(';').strip
		cdm_data['Contributor (Corporate)'] = record.grab('710').map(&:value).join(';').strip
		cdm_data['Contributor (Conference or Meeting)'] = record.grab('711').map(&:value).join(';').strip
		cdm_data['Personal name added/Title'] = record.find_all {|f| f.tag == '700' and f['t']}.map(&:value).join(';').strip
		cdm_data['Related Resource (Serial Title)'] = record['730'].value rescue nil
		cdm_data['Related Publication'] = record['787'].value rescue nil
		cdm_data['Series (controlled)'] = record.grab('8[03]0').map(&:value).join(';').strip
		cdm_data['Electronic access'] = record.find_all {|f| f.tag == '856' and f['u'] !~ /\.jpg/}.map{|f| f['u']}.join(';').strip
		cdm_data['Filename (verified)'] = record.find_all {|f| f.tag == '856' and f['3']}.map{|f| f['3'].scan(/\d+\.jpg/)}.join(';').strip
		cdm_data['Call # (without location prefix)'] = record.find_all {|f| f.tag == '856' and f['3']}.map{|f| f['3'].split(';')[0]}.join(';').strip
		cdm_data['Shelving Location of Physical Item(s)'] = record.grab('949').map{ |f| f['d'] }.join(';').strip
		cdm_data['Horizon bib #'] = record['996'].value rescue nil
		cdm_data['Object Type'] = record.type
		cdm_data['OCLC#'] = record.oclc_number
		cdm_data['Digitization Note'] = record.grab_by_value('500', keywords[:digitization]).map(&:value).join(';').strip
		cdm_data['Unidentified notes'] = record.find_all {|f| f.tag == '500' and f.value !~ /#{note_string}/i}.map(&:value).join(';').strip

		cdm_data.each { |k, v| cdm_data[k] = ' ' if v.nil? or v.empty? }
		raise "Consistency error: keys are not equal." unless cdm_data.keys == spec.keys
		cdm << cdm_data
	rescue Exception => ex
		puts record['001'].value + ' -- ' + ex.backtrace.join("\n")
	end
end

puts "RECORDS READ: " + num_records.to_s
cdm.sort_by! { |cdm_data| cdm_data[:title] }

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
cdm.each_with_index { |h, idx| idx += 1; transform.row(idx).replace(h.values) }
book.write "cdm-#{Time.now.strftime('%Y%m%d')}.xls"