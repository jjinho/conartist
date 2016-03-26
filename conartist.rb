#!/usr/bin/ruby

#1 Get rid of all "Header Blocks"
#  First Header Block starts with ^R^O
#  First Header Block ends with 130 '-'
#2 Get rid of all subsequent "Header Blocks"
#  Other Header Blocks start with ^L^R^O
#  Other Header Blocks end with 130 '-'
#3 Get Total Number of Cases
#  Match line "Total # Cases "
#  Number at the end is the sought after value
#4 Get rid of "Footer Block"
#  Footer Block starts from 30 '*'
#  Footer Block ends at EOF
#

require 'axlsx'
require 'titleize'

def conartist (input_filename, output_filename)
  rpt = []
  rpt_clean = []
  
  f = File.open(input_filename, 'r')
  
  # Getting rid of Header Blocks
  skip_header = false
  f.each do |line|
    if line =~ /^\u0012/ or line =~ /^\u000C/
      skip_header = true
    end
    if line =~ /^-+/
      skip_header = false
      next
    end
    if not skip_header
      rpt.push(line)
    end
  end
  
  f.close()
  
  # Geting the total number of cases
  num_cases = 0
  rpt.each do |line|
    if line =~ /Total # Cases/
      num_cases = line.scan(/\d+/).first.to_i
    end
  end
  
  # Getting rid of Footer Block
  skip_footer = false
  rpt.each do |line|
    if line =~ /^\*+/
      skip_footer = true
    end
    if not skip_footer
      rpt_clean.push(line)
    end
  end
  
  #puts rpt_clean
  
  or_cases = []
  # Get each case
  or_case_total = {}
  rpt_clean.each do |line|
    or_case = {}
    # Make sure we're not dealing with a blank line
    if line.strip.length > 0
      or_case[:mrn]       = line[0, 10].strip
      or_case[:date]      = line[10, 10].strip
      or_case[:age]       = line[20, 4].strip
      or_case[:sex]       = line[24, 4].strip
      or_case[:preop]     = line[28, 20].strip
      or_case[:postop]    = line[48, 10].strip
      or_case[:proc]      = line[58, 22].strip
      or_case[:surgeon]   = line[80, 14].strip
      or_case[:cosurgeon] = line[94, 12].strip
      or_case[:assist]    = line[106, line.length - 106].strip
      # Is this the first line?
      if or_case[:mrn].length > 0 or
         or_case[:date].length > 0 or
         or_case[:age].length > 0 or
         or_case[:sex].length > 0
        # This is a first line
        # Store the previous or_case_total
        or_cases.push(or_case_total)
  
        # Set the new or_case_total
        or_case_total = or_case
      else
        # This is not a first line
        or_case_total[:preop]     = "#{or_case_total[:preop]} #{or_case[:preop]}".rstrip
        or_case_total[:postop]    = "#{or_case_total[:postop]} #{or_case[:postop]}".rstrip
        or_case_total[:proc]      = "#{or_case_total[:proc]} #{or_case[:proc]}".rstrip
        or_case_total[:surgeon]   = "#{or_case_total[:surgeon]} #{or_case[:surgeon]}".rstrip
        or_case_total[:cosurgeon] = "#{or_case_total[:cosurgeon]} #{or_case[:cosurgeon]}".rstrip
        or_case_total[:assist]    = "#{or_case_total[:assist]} #{or_case[:assist]}".rstrip
      end
    end
  end
  # Add the last case
  or_cases.push(or_case_total)
  # Get rid of the empty first case
  or_cases.slice!(0)
  
  # Create the Excel sheet
  p = Axlsx::Package.new
  wb = p.workbook
  wb.styles do |s|
    # Excel sheet styles
    header_style = s.add_style :sz => 26,
                               :alignment => { :horizontal => :center,
                                               :vertical => :center }
    subheader_style = s.add_style :sz => 18,
                                  :alignment => { :horizontal => :center,
                                                  :vertical => :center }
    title_style = s.add_style :sz => 16, 
                              :b => true,
                              :alignment => { :horizontal => :left,
                                              :vertical => :center },
                              :border => { :style => :thin,
                                           :color => "000000",
                                           :edges => [:top] }
    data_style = s.add_style :sz => 16,
                             :alignment => { :horizontal => :left,
                                             :vertical => :top,
                                             :wrap_text => true },
                              :border => { :style => :thin, 
                                           :color => "000000", 
                                           :edges => [:top, :bottom] }
    # Excel worksheets
    wb.add_worksheet(:name => "Operative Reprt") do |sheet|
      # Headers
      sheet.add_row ["Richmond University Medical Center", "", "", "", "", "", "", "", "", ""], :style => header_style
      sheet.merge_cells("A1:J1")
      sheet.add_row ["Department of Surgery Morbidity and Mortality Conference", "", "", "", "", "", "", "", "", ""], :style => subheader_style
      sheet.merge_cells("A2:J2")
      sheet.add_row ["#{or_cases.first[:date]} - #{or_cases.last[:date]}", "", "", "", "", "", "", "", "", ""], :style => subheader_style
      sheet.merge_cells("A3:J3")
      sheet.add_row ["", "", "", "", "", "", "", "", "", ""]
      sheet.merge_cells("A4:J4")
      sheet.add_row ["MRN", 
                     "DOS", 
                     "Age", 
                     "Sex", 
                     "Preoperative Diagnosis", 
                     "Postoperative Diagnosis", 
                     "Procedure", 
                     "Pathology", 
                     "Surgeon", 
                     "Assistant"], :style => title_style
      or_cases.each do |or_case|
        sheet.add_row [or_case[:mrn],
                       or_case[:date],
                       or_case[:age],
                       or_case[:sex],
                       or_case[:preop].capitalize,
                       or_case[:postop].capitalize,
                       or_case[:proc].capitalize,
                       "",
                       or_case[:surgeon].titleize,
                       or_case[:assist].titleize], :style => data_style
      end
      sheet.column_widths 14, 13.5, 9.5, 7.5, 43, 34.5, 63, 44, 17.5, 15.5
      sheet.page_setup.fit_to :width => 1, :orientation => :landscape
    end
  end
  
  p.serialize output_filename
end

conartist(ARGV[0], ARGV[1])