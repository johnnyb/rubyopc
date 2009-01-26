require "tempfile"
require "rexml/document"
require "rexml/streamlistener"

class String
  #FIXME - only define if doesn't exist
  def blank?
    !self.match(/^[\s]*$/).nil?
  end
end


module OpenPackagingConventions
  module Docx
    class SimpleWriter
      def initialize
        @style_list = []
        @list_list = []

        @doc = REXML::Document.new
        @doc << REXML::XMLDecl.new("1.0", "UTF-8", "yes")
        doc = @doc.add_element("w:document", "xmlns:w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
        body = doc.add_element("w:body")

        @context_stack = []
        # Context:
        # @table_point
        # @text_insertion_point
        # @paragraph_insertion_point
        # @state
        @table_point = nil
        @text_insertion_point = nil
        @paragraph_insertion_point = body
        @state = :begin
      end

      def push_context
        @context_stack.push({
                              :table_point => @table_point, 
                              :text_insertion_point => @text_insertion_point, 
                              :paragraph_insertion_point => @paragraph_insertion_point, 
                              :state => @state
                            })
      end

      def pop_context
        ctx = @context_stack.pop
        @paragraph_insertion_point = ctx[:paragraph_insertion_point]
        @text_insertion_point = ctx[:text_insertion_point]
        @table_point = ctx[:table_point]
        @state = ctx[:state]
      end

      ### INTERNAL FUNCTIONS - should I move these to a base class?
      def to_docx(fname)
        to_opc(fname)
      end

      def to_opc(fname)
        OpenPackagingConventions::Package.with_package(fname) do |p|
          p.add_part("/word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", document_xml)
          p.add_part_to("/word/document.xml", "/word/styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", style_xml)
          p.add_part_to("/word/document.xml", "/word/numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", numbering_xml)
        end
      end

      def close_paragraph
        @text_insertion_point = nil

        @state = :closed_paragraph
      end

      def register_style(style_info)
      end

      def register_list(list_info)
        list_num_id = @list_list.size + 1
        @list_list.push(list_info.merge(:numId => list_num_id))
        return list_num_id
      end

      def open_paragraph(style_info)
        w_p = @paragraph_insertion_point.add_element("w:p")

        # FIXME - generate paragraph styles
        paragraph_style = style_info[:pStyle]

        w_ppr = w_p.add_element("w:pPr")
        w_ppr.add_element("w:pStyle", "w:val" => paragraph_style)

        unless style_info[:numPr].nil?
          w_numpr = w_ppr.add_element("w:numPr")
          w_numpr.add_element("w:ilvl", "w:val" => style_info[:numPr][:ilvl])
          w_numpr.add_element("w:numId", "w:val" => style_info[:numPr][:numId])
        end

        @text_insertion_point = w_p
        @state = :paragraph
      end
        
      def write_run(txt, style_info)
        open_paragraph(style_info) unless @text_insertion_point.nil?
        w_r = @text_insertion_point.add_element("w:r")
        w_rpr = w_r.add_element("w:rPr")

        # FIXME - generate run styles 
        run_style = style_info[:run_style] || "Normal"

        w_rpr.add_element("w:rStyle", "w:val" => run_style)

        w_t = w_r.add_element("w:t")
        w_t.add_text(txt)
      end

      def open_table(style_info)
        close_paragraph
        push_context

        insertion_point = @paragraph_insertion_point
        
        w_tbl = insertion_point.add_element("w:tbl")
        w_tblpr = w_tbl.add_element("w:tblPr")
        w_tblborders = w_tblpr.add_element("w:tblBorders")
        ["w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"].each do |prop|
          w_tblborders.add_element(prop, "w:val" => "single", "w:sz" => "1")
        end
        w_tblpr.add_element("w:tblStyle", "w:val" => "TableGrid")
        w_tblpr.add_element("w:tblW", "w:w" => "0", "w:type" => "auto")
        # w_tblpr.add_element("w:tblLook", "w:val" => "04A0")

        w_tbl.add_element("w:tblGrid")

        @table_point = w_tbl
        @text_insertion_point = nil
        @paragraph_insertion_point = nil
        @state = :table_begin
      end

      def close_table
        col_info = []
        @table_point.elements.each do |tr_e|
          if tr_e.name == "tr"
            1.upto(tr_e.elements.size) do |tc_idx|
              tc_e = tr_e.elements[tc_idx]
              if tc_e.name == "tc"
                col_info[tc_idx - 1] = {}
                #FIXME - search for <w:tcW w:w="" /> node and use that for width
                #widths are specified in 1/20 of a point
              end
            end
          end
        end
       
        w_tblgrid = @table_point.elements[2]
        col_info.each do |cinfo|
          w_tblgrid.add_element("w:gridCol", "w:w" => "1024")
        end
        
        pop_context
      end

      # Note - style info is ignored for rows
      def open_table_row(style_info)
        w_tr = @table_point.add_element("w:tr")
      end

      def close_table_row
      end

      def open_table_cell(style_info)
        
        w_tr = @table_point.elements[@table_point.elements.size]
        raise "now rows found! #{w_tr.name}" unless w_tr.name == 'tr'

        w_tc = w_tr.add_element("w:tc")
        w_tcpr = w_tc.add_element("w:tcPr")

        # FIXME - need better way of calculating width
        w_tcpr.add_element("w:tcW", "w:w" => "1024")

        @paragraph_insertion_point = w_tc
        @state = :begin_table_cell
      end 

      def close_table_cell
      end

      # Return style information for this document
      def style_xml
        #FIXME - return two abstracts - ul/ol, and then one instance per.
      
        return <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="normal">
    <w:name w:val="normal" />
    <w:rPr>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="dt">
    <w:name w:val="Item Heading" />
    <w:rPr>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="dd">
    <w:name w:val="Item Content" />
    <w:rPr>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="p">
    <w:name w:val="General Paragraph" />
    <w:rPr>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="h1">
    <w:name w:val="Heading Level 1" />
    <w:rPr>
      <w:b />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="h2">
    <w:name w:val="Heading Level 2" />
    <w:rPr>
      <w:b />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="h3">
    <w:name w:val="Heading Level 3" />
    <w:rPr>
      <w:b />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="h4">
    <w:name w:val="Heading Level 4" />
    <w:rPr>
      <w:b />
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ol_li">
    <w:name w:val="Ordered List Item" />
    <w:rPr>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ul_li">
    <w:name w:val="Unordered List Item" />
    <w:rPr>
    </w:rPr>
  </w:style>
</w:styles>
EOF

      end

      # Return numbering information for this document
      def numbering_xml
        # 0 is numbered lists, 
        # 1 is bulleted lists

        doc = REXML::Document.new
        doc << REXML::XMLDecl.new("1.0", "UTF-8", "yes")

        w_numbering = doc.add_element("w:numbering", "xmlns:w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        # Numbered Lists
        w_abstractnum = w_numbering.add_element("w:abstractNum", "w:abstractNumId" => "0")
        w_abstractnum.add_element("w:multiLevelType", "w:val" => "hybridMultiLevel")
        w_lvl = w_abstractnum.add_element("w:lvl", "w:ilvl" => "0")
        w_lvl.add_element("w:start", "w:val" => "1")
        w_lvl.add_element("w:numFmt", "w:val" => "decimal")
        w_lvl.add_element("w:lvlText", "w:val" => "%1.")
        w_lvl.add_element("w:lvlJc", "w:val" => "left")
        w_ppr = w_lvl.add_element("w:pPr")
        w_ppr.add_element("w:ind", "w:left" => "720", "w:hanging" => "360")

        # Bulletted Lists
        w_abstractnum = w_numbering.add_element("w:abstractNum", "w:abstractNumId" => "1")
        w_abstractnum.add_element("w:multiLevelType", "w:val" => "hybridMultiLevel")
        w_lvl = w_abstractnum.add_element("w:lvl", "w:ilvl" => "0")
        w_lvl.add_element("w:start", "w:val" => "1")
        w_lvl.add_element("w:numFmt", "w:val" => "bullet")
        w_lvl.add_element("w:lvlText", "w:val" => "&#239;&#8218;&#183;")
        w_lvl.add_element("w:lvlJc", "w:val" => "left")
        w_ppr = w_lvl.add_element("w:pPr")
        w_ppr.add_element("w:ind", "w:left" => "720", "w:hanging" => "360")
        w_rpr = w_lvl.add_element("w:rPr")
        w_rpr.add_element("w:rFonts", "w:ascii" => "Symbol", "w:hAnsi" => "Symbol", "w:hint" => "default")

        @list_list.each do |l_info|
          w_num = w_numbering.add_element("w:num", "w:numId" => l_info[:numId])
          w_num.add_element("w:abstractNumId", "w:val" => (l_info[:type] == :ordered_list ? 0 : 1))
        end

        output = ""
        doc.write output
        return doc
      end

      # Return main document text
      def document_xml
        output = ""
        @doc.write output
        return output
      end

    end

    class XMLConversionWriter < SimpleWriter
      include REXML::StreamListener

      attr_accessor :tag_map

      attr_accessor :stylesheet  #Not yet used

      def initialize
        @style_stack = []
        @list_stack = []

        self.tag_map = DEFAULT_TAG_MAP

        super
      end

      def parse_stream(data)
        doc = REXML::Document.parse_stream(data, self)

        return self
      end

      # This adds style info into the stack and adds additional interpretive stuff to it
      def push_style_info(current_style)
        tag = current_style[:tag]
        tag_type = tag_map[tag]
        current_style[:run_style] = tag
        current_style[:paragraph_style] = tag
        current_style[:tag_type] = tag_type

        if tag_type == :unordered_list || tag_type == :ordered_list
          previous_levels = @style_stack.select{|s| s[:tag_type] == tag_type}
          if previous_levels.empty?
            current_style[:list_numId] = register_list({:type => tag_type})
            current_style[:list_ilvl] = 0
          else
            current_style[:list_numId] = previous_levels.last[:list_numId]
            current_style[:list_ilvl] = previous_levels.last[:list_ilvl] + 1
          end
        end
        
        if tag_type == :list_item
          list_parents = @style_stack.select{|s| s[:tag_type] == :ordered_list || s[:tag_type] == :unordered_list}
          unless list_parents.empty?
            current_style[:list_numId] = list_parents.last[:list_numId]
            current_style[:list_ilvl] = list_parents.last[:list_ilvl]
          else
            #If there is no ul/ol parent, just treat like a paragraph
          end
        end

        @style_stack.push(current_style)
      end

      def word_style_info_for(style_stack = nil)
        # FIXME - just a temporary hack to get things started
        # Also need to register styles if appropriate
        # Should eventually be based on CSS
        style_stack = @style_stack if style_stack.nil?
        current_style = style_stack.last
        word_style = {}

        word_style[:action] = current_style[:tag_type]
        word_style[:pStyle] = current_style[:paragraph_style]
        word_style[:rStyle] = current_style[:run_style]
        if current_style[:tag_type] == :list_item
          word_style[:numPr] = { :ilvl => current_style[:list_ilvl], :numId => current_style[:list_numId]}
        end

        return word_style
      end

      def write_text(txt)
        return if @state == :ignore
        word_style_info = word_style_info_for(@style_stack)
        open_paragraph(word_style_info) unless @state == :paragraph
        write_run(txt, word_style_info)
      end

      def open_block(style_info)
        close_paragraph if @state == :paragraph
      end

      def close_block
        close_paragraph if @state == :paragraph
      end

      def open_inline(style_info)
      end

      def close_inline
      end

      def open_ignore(style_info)
        @before_ignore_state = @state
        @state = :ignore
      end

      def close_ignore
        @state = @before_ignore_state
      end
      
      def cdata(content)
        text(content)
      end

      @@entities = [
                    ["&nbsp;", "&#160;"],
                    ["&Acirc;", "&#194;"],
                    ["&sect;", "&#167;"] ,
                    ["&Atilde;", "&#195;"]
                   ]
      
      def substitute_entities(str)
        @@entities.each do |ent|
          str = str.gsub(*ent)
        end
        str = str.gsub(/\&([^#])/, "&amp;\\1")
        return str
      end
        
      def text(content)
        content = substitute_entities(content)
        content = content.gsub("\n", " ")
        # puts "GENERATING TEXT: #{content}"
        write_text(content) unless content.blank?
      end
      
      def tag_start(name, attrs)
        # Parse out the tags/classes/ids into symbols
        name = name.to_sym

        # puts "TAG START: #{name}"

        # Get style information
        classes = (attrs["class"]||"").split(/\s+/).map{|x| x.to_sym}
        tag_id = attrs["id"]
        tag_id = tag_id.to_sym unless tag_id.nil?
        style_info = { :tag => name, :classes => classes, :id => tag_id }
        push_style_info(style_info)
        word_style_info = word_style_info_for(@style_stack)
        tag_type = tag_map[name]
        
        # Determine function to call
        func = tag_type
        func = :block if tag_type == :list_item

        # Skip these
        return if [:ordered_list, :unordered_list, :skip, :ignore].include?(tag_type)

        # Call appropriate function
        unless func.nil?
          self.send("open_#{func}", word_style_info)
        else
          self.open_inline(word_style_info)
        end
      end

      def tag_end(name)
        @style_stack.pop
        
        name = name.to_sym
        tag_type = tag_map[name]

        # Determine function to call
        func = tag_type
        func = :block if tag_type == :list_item

        # Skip these
        return if [:ordered_list, :unordered_list, :skip].include?(tag_type)

        # Class appropriate function
        unless func.nil?
          self.send("close_#{func}")
        else
          self.close_inline
        end
          
      end


      #FIXME - eventually this will be replaced by looking the tag up in the stylesheet
      DEFAULT_TAG_MAP = {
        :h1 => :block,
        :h2 => :block,
        :h3 => :block,
        :h4 => :block,
        :h5 => :block,
        :p => :block,
        :div => :block,
        
        :i => :inline,
        :em => :inline,
        :strong => :inline,
        :b => :inline,
        :u => :inline,
        
        :script => :ignore,
        :style => :ignore,
        :title => :ignore,
        
        :ol => :ordered_list,
        :ul => :unordered_list,

        :li => :list_item,
        
        :table => :table,
        :tr => :table_row,
        :td => :table_cell,
        :th => :table_cell,
        
        :tbody => :skip
          
      }
    end


    module Converter
      class HTML
        def initialize
          @docx_writer = OpenPackagingConventions::Docx::XMLConversionWriter.new
        end

        def docx_writer
          @docx_writer
        end

        def self.docx_writer_from_html_data(hdata)
          w = OpenPackagingConventions::Docx::XMLConversionWriter.new
          w.parse_stream(hdata)
          return w
        end

        def self.docx_from_html_data(html_data, docx_fname, opts = {})
          begin
            t = Tempfile.new("html")
            f = File.open(t.path, "w")
            f.write(html_data)
            f.close
            docx_from_html(t.path, docx_fname, opts)
          ensure
            t.unlink
          end
        end

        def self.docx_from_html(html_fname, docx_fname, opts = {})
          if opts[:tidy_html]
            orig_html_fname = html_fname
            t = Tempfile.new("html")
            html_fname = t.path
            t.close
            t.unlink
            system("tidy", "-o", html_fname, "-q", "-asxhtml", orig_html_fname)
          end
          
          html_data = nil
          File.open(html_fname){|f| html_data = f.read}
          docx_writer = self.docx_writer_from_html_data(html_data)
          opc = docx_writer.to_opc(docx_fname)
          if opts[:tidy_html]
            File.unlink(html_fname)
          end
        end
      end
    end
  end
end

