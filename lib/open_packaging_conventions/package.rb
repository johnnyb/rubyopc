require "tempfile"
module OpenPackagingConventions

  class Package
    def initialize(fname = nil)
      @relationships = {}
      @documents = {}
      @content_types = {}
      @id_seq = 0
      @path_ids = {}
      @top_dir = get_tmpdir
      @file_name = fname
    end

    # Scoped block for use with resource management.
    # If you give it a filename, it will attempt to save it unless an error is thrown.
    # Using this format ensures that close() is called, which will clean up any directory problems.
    def self.with_package(fname = nil, &block)
      p = Package.new(fname)
      begin
        yield(p)
        p.save unless fname.nil?
      ensure
        p.close
      end
    end

    def close
      remove_tmpdir
    end

    # Returns the id for a given path
    def id_for_path(path)
      if @path_ids[path].nil?
        @path_ids[path] = "id#{@id_seq}"
        @id_seq = @id_seq + 1
      end
      return @path_ids[path]
    end

    # Adds a toplevel part to the package
    # - path is the path in the package for the file
    # - relationship is the URL specifying the relationship
    # - content_type is the content-type of the file
    # - data is the full contents of the file
    def add_part(path, relationship, content_type, data)
      add_part_to("/", path, relationship, content_type, data)
    end

    # Adds a subpart to the package
    # - base_path is the path to the file that this is being attached to
    # - path is the path in the package for the current file being added
    # - relationship is the URL specifying the relationship
    # - content_type is the content-type of the file
    # - data is the full contents of the file
    def add_part_to(base_path, path, relationship, content_type, data)
      raise "InvalidPath" unless path[0..0] == "/"
      raise "InvalidPath" unless path.index("/../").nil?

      @relationships[base_path] ||= []
      @relationships[base_path].push([path, relationship])
      @content_types[path] = content_type
      @documents[path] = true
      make_directory_for_path(path)
      File.open("#{@top_dir}#{path}", "w"){|f| f.write(data)}
    end

    def save
      save_as(@file_name)
    end

    # Outputs the package to a given file name
    def save_as(fname)
      rewrite_metadata

      curwd = Dir.getwd
      Dir.chdir @top_dir
      tmp_out = Tempfile.new("docxzip")
      tmp_out_path = tmp_out.path
      tmp_out.close
      tmp_out.unlink

      #FIXME - need to redirect stdout/stderr
      system("zip", "-q", "-r", tmp_out_path, ".")
      Dir.chdir(curwd)
      File.rename(tmp_out_path, fname)
    end



    private

    #FIXME - should use Builder for this
    def content_types_file
      return <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  #{@content_types.map{|key, val| "<Override PartName='" + key + "' ContentType='" + val + "' />"}.join("\n  ")}
</Types>
EOF
  end

  def rels_for(path)
    return nil if @relationships[path].nil?

    relationships_str = @relationships[path].map{|info|
      subpath = info[0]
      relationship = info[1]
      tag_id = id_for_path(subpath)
      relative_path = relative_path(path, subpath)
      "<Relationship Id='#{tag_id}' Type='#{relationship}' Target='#{relative_path}' />"
    }.join("\n  ")
    
    return <<EOF
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  #{relationships_str}
</Relationships>
EOF
  end


  # Makes sure that a directory gets made.
  # - path is the path for which we need to make directories
  def make_directory_for_path(path)
    comps = path.split("/")
    comps.pop

    1.upto(comps.size - 1) do |comp_idx|
      dirname = comps[0..comp_idx].join("/")
      full_path = "#{@top_dir}#{dirname}"
      Dir.mkdir(full_path) unless File.exists?(full_path)
    end
  end

  # Retrieves the path to a temporary directory.  Grabs a tempfile, then deletes it and creates a directory with the same name.
  def get_tmpdir
    path = nil
    t = Tempfile.open("docx") { |f|
      path = f.path
    }
    File.unlink(path)
    Dir.mkdir(path)

    return path
  end

  def remove_tmpdir
    system("rm", "-R", @top_dir)
  end

  def rewrite_metadata
    # Write top-level listing of content types
    path = "/[Content_Types].xml"
    File.open("#{@top_dir}#{path}", "w"){|f| f.write(content_types_file)}

    # Write top-level relationships
    path = relationship_path("/")
    make_directory_for_path(path)
    File.open("#{@top_dir}#{path}", "w"){|f| f.write(rels_for("/"))}

    # Write sub-relationships
    @documents.keys.each do |path|
      unless @relationships[path].nil?
        rel_path = relationship_path(path)
        make_directory_for_path(rel_path)
        File.open("#{@top_dir}#{rel_path}", "w"){|f| f.write(rels_for(path))}
      end
    end
  end

  # FIXME - not yet implemented
  def load_metadata
  end

  def relationship_path(path)
    if path == "/"
      # For some reason I couldn't get / to work, so I'm just going to special-case it
      return "/_rels/.rels"
    end
    comps = path.split("/")
    comps = [""] if comps.empty?
    fname = comps.pop + ".rels"
    comps.push("_rels")
    comps.push(fname)
    return comps.join("/")
  end

  def relative_path(start, finish)
    start_comps = start.split("/")
    end_comps = finish.split("/")

    # FIXME - TEMPORARY HACK - need to fix the relative_path stuff
    return start == "/" ? finish[1..-1] : end_comps[-1]

    num_similar = 0
    should_stop = false
    start_comps.each_index do |comp_idx|
      unless should_stop  
        if start_comps[comp_idx] == end_comps[comp_idx]
          num_similar = num_similar + 1
        else
          should_stop = true
        end
      end
    end
    
    backpaths = (end_comps.size - start_comps.size) + (end_comps.size - num_similar) - 1
    final_path_comps = (1..backpaths).map{|x| ".."} + end_comps[-backpaths..-1]
    return final_path_comps.join("/")
  end
end



end
