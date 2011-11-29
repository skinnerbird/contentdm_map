require 'scl'

# SCL Metadata Application Profile processor
# Generates .xls in accordance with specification
# for import into Contentdm
# Original: 09/19/2011, Updated: 11/28/2011
# Author: Mark Cooper

=begin
SET CRITERIA / CONSTRAINTS FOR FIELDS
=end

NUM_HEADERS = 65
spec = {}
File.open('spec.txt').each_line do |l|
	header, width = l.split('|')
	spec[header] = width.strip
end

raise "Incorrect number of headers" unless spec.size == NUM_HEADERS

num_records     = 0
transform       = {}
headers         = {}
cdm             = []
transform[:cdm] = cdm
headers[:cdm]   = spec

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
CREATE THE SPREADSHEET
=end

MARC::ForgivingReader.new('PhotoRecordConversions_CDM_ready_TEST.mrc').each do |r|
	num_records += 1
	record = MarcTools::MarcRecordWrapper.new(r)
	cdm_data = Hash.new
	cdm_data['Creator'] = record['100'] ? MarcTools::MarcFieldWrapper.new(record['100']).subfield_string(' ', 'e').strip : ''
	cdm_data['Creator Role'] = record['100'] ? record['100']['e'] : ''
	cdm_data['Title'] = MarcTools::MarcFieldWrapper.new(record['245']).subfield_string(' ', 'c', 'h').strip
	cdm_data['Statement of Responsibility'] = record['245']['c'] # REC MUST HAVE 245
	cdm_data['Alternative Title'] = record.grab('246').map(&:value).join(';').strip
	cdm_data['Place of Publication (Original)'] = record['260'] ? record['260']['a'] : '' # MUST HAVE 260
	cdm_data['Publisher (Original)'] = record['260'] ? record['260']['b'] : ''
	cdm_data['Creation/Publication Date'] = record['260'] ? record['260']['c'] : '' # date1 ???
	cdm_data['Copyright Date'] = record.copyright_date
	cdm_data['Reprint Date'] = record.dtst == 'r' ? record.date1 : ''
	cdm_data['Creation/Publication Date (clean)'] = record.dtst == 'r' ? record.date1 : ''
	cdm_data['Creation/Publication Date (2) (clean)'] = record.dtst == 'r' ? record.date2 : ''
	cdm_data['Object Description'] = record.grab('520').map(&:value).join(';').strip
	cdm_data['Biographical/Historical Note'] = record['545'] ? record['545'].value : ''
	cdm_data['Table Of Contents '] = record.grab('505').map(&:value).join(';').strip
	cdm_data['Creator Note'] = record.grab_by_value('500', keywords[:creator_note]).map(&:value).join(';').strip
	cdm_data['Number of Images/Parts'] = record['300'] ? record['300']['a'].gsub(/:\s*$/, '').strip : ''
	cdm_data['Other Physical Details'] = record['300']['b'] ? record['300']['b'].gsub(/;\s*$/, '').strip : ''
	cdm_data['Dimensions'] = record['300']['c'] # REC MUST HAVE 300 FOR THIS & ABOVE OR ERROR
	cdm_data['Language'] = record.grab('041').map(&:subfields).join(';').strip # TEST
	cdm_data['Source of Title'] = record.grab_by_value('500', keywords[:source_of_title]).map(&:value).join(';').strip
	cdm_data['Source of Title Variations'] = record.grab_by_value('500', keywords[:source_of_alt_title]).map(&:value).join(';').strip
	cdm_data['Original Creation/Publication Details'] = record.grab_by_value('500', keywords[:publication_note]).map(&:value).join(';').strip
	cdm_data['Reproduction Details (physical item)'] = record.grab_by_value('500', keywords[:reproduction]).map(&:value).join(';').strip
	cdm_data['Location of Originals'] = record['535'] ? record['535'].value : ''
	cdm_data['Collection Guide'] = 'SUPPLIED' # TEMPLATE
	cdm_data['Owning Institution'] = record.grab_by_value('500', keywords[:ownership]).map(&:value).join(';').strip
	cdm_data['Copyright Status'] = record['540'] ? record['540']['a'] : ''
	cdm_data['Reuse and Reproduction Restrictions'] = 'SUPPLIED' # TEMPLATE
	cdm_data['Immediate Source of Acquisition'] = record['541'] ? record['541'].value : ''
	cdm_data['Provenance'] = record['561'] ? record['561'].value : ''
	cdm_data['Physical Condition'] = record.grab_by_value('500', keywords[:condition]).map(&:value).join(';').strip
	cdm_data['Signatures and Inscriptions'] = record.grab_by_value('500', keywords[:signatures]).map(&:value).join(';').strip
	cdm_data['Numbers or Letters on Originals'] = record.grab_by_value('500', keywords[:numbering]).map(&:value).join(';').strip
	cdm_data['Markings and Stamps'] = record.grab_by_value('500', keywords[:marks_stamps]).map(&:value).join(';').strip
	cdm_data['Other Statements of Responsibility'] = 'N/A'
	cdm_data['Source Note'] = record.grab_by_value('500', keywords[:source_note]).map(&:value).join(';').strip
	cdm_data['Local Note'] = record['590'] ? record['590'].value : ''
	cdm_data['Subject (Person)'] = record.find_all {|f| f.tag == '600' and f.indicator1 =~ /(0|1)/}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Family)'] = record.find_all {|f| f.tag == '600' and f.indicator1 == '3'}.map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Corporate Body)'] = record.grab('610').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Meeting/Event)'] = record.grab('611').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Time Period)'] = record.grab_by_value('648', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Topical)'] = record.grab_by_value('650', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip
	cdm_data['Subject (Geographic FAST)'] = record.grab_by_value('651', 'fast$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(', ', '2').strip.chop}.join(';').strip
	cdm_data['Subject (Geographic local)'] = record.grab_by_value('651', 'local$').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(', ', '2').strip.chop}.join(';').strip
	cdm_data['Street Address'] = record['939'] ? record['939']['a'] : ''
	cdm_data['Location Coordinates'] = record['938'] ? record['938']['a'] : ''
	cdm_data['Genre/Form'] = record.grab('655').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', '2').strip}.join(';').strip	
	cdm_data['Contributor (Person)'] = record.grab('700').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', 'e', 't').strip}.join(';').strip
	cdm_data['Contributor (Corporate)'] = record.grab('710').map(&:value).join(';').strip
	cdm_data['Contributor (Conference or Meeting)'] = record.grab('711').map(&:value).join(';').strip
	cdm_data['Related Resource'] = record.find_all {|f| f.tag == '700' and f['t']}.map(&:value).join(';').strip
	cdm_data['Related Resource (Uniform Title)'] = record['730'] ? record['730'].value : ''
	cdm_data['Related Publication'] = record['787'] ? record['787'].value : ''
	cdm_data['Collection'] = record.grab('8[03]0').map(&:value).join(';').strip
	cdm_data['URI'] = record.find_all {|f| f.tag == '856' and f['u'] !~ /\.jpg/}.map{|f| f['u']}.join(';').strip
	cdm_data['Archival Disc File ID'] = record.find_all {|f| f.tag == '856' and f['3']}.map{|f| f['3'].scan(/\d+\.jpg/)}.join(';').strip
	cdm_data['Call Number'] = record.find_all {|f| f.tag == '856' and f['3']}.map{|f| f['3'].split(';')[0]}.join(';').strip
	cdm_data['Shelving Location of Physical Item(s)'] = record.grab('949').map{ |f| MarcTools::MarcFieldWrapper.new(f).subfield_string(' ', 'a', 'b', 'c', 'm', 'q', 't').strip}.join(';').strip
	cdm_data['Identifier'] = record['996'] ? record['996'].value : ''
	cdm_data['Object Type'] = record.type
	cdm_data['OCLC#'] = record.oclc_number
	cdm_data['Description'] = record.grab_by_value('500', keywords[:digitization]).map(&:value).join(';').strip
	cdm_data['Unidentified notes'] = record.find_all {|f| f.tag == '500' and f.value !~ /#{note_string}/i}.map(&:value).join(';').strip

	cdm_data.each { |k, v| cdm_data[k] = ' ' if v.nil? or v.empty? }
	raise "Consistency error - data chunk for|| " + record['245']['a'] + "|| does not match headers size." unless cdm_data.size == NUM_HEADERS
	cdm << cdm_data
end

puts "RECORDS READ: " + num_records.to_s
cdm.sort_by! { |cdm_data| cdm_data[:title] }

SCL::excel_report transform, headers, "cdm-#{Time.now.strftime('%Y%m%d')}.xls"

# ex = SCL::ExcelReport.new(cdm)
# ex.name 'Contentdm'
# ex.header(spec.keys)
# ex.fit(spec.values)
# format = {
# 	:weight           => :bold,
# 	:color            => :white,
# 	:pattern_fg_color => :blue, 
# 	:pattern          => 1,
# 	:horizontal_align => :center,
# 	:vertical_align   => :center,
# }
# ex.fix_height(0, 18)
# ex.apply_format(0, format)
# ex.process
# ex.print("cdm-#{Time.now.strftime('%Y%m%d')}.xls")