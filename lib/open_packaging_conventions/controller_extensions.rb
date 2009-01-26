module OpenPackagingConventions

  # This module mixes in to ApplicationController to provide helper methods
  # for creating docx files
  module ControllerExtensions

    # Takes the html data and a filename and produces a docx file and
    # sends it to the client.  Use this _instead_ of a render or send
    # command in your controller
    def render_html_as_docx(html_data, fname)
      begin
        t = Tempfile.new("docx")
        t.close
        OpenPackagingConventions::Docx::Converter::HTML.docx_from_html_data(html_data, t.path, :tidy_html => true)
        File.open(t.path) {|f|
          send_data(f.read, :filename => fname)
        }
      ensure
        t.close
        t.unlink
      end
    end
  end
end
