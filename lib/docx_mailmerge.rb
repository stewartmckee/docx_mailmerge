require 'set'
require 'zip'
require 'nokogiri'
require 'docx_mailmerge/version'

module DocxMailmerge
  class MailMerge
    attr_accessor :zip, :parts
    def initialize(f)
      if f.respond_to?(:read)
        @zip = Zip::File.new(f, buffer = true)
      elsif f.class == String
        @zip = Zip::File.open(f)
      else
        raise ArgumentError, 'File must be either a filename or an IO-like object.'
      end
      @parts = {}

      # Figure out which files we need and open them up
      relevant_content_types = [
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      ].map { |x| "xmlns:Types/xmlns:Override[@ContentType='#{x}']" }.join(' | ')
      zip_read_xml('[Content_Types].xml').xpath(relevant_content_types).each do |f|
        fn = f['PartName'].split('/', 2)[1]
        @parts[fn] = zip_read_xml(fn)
      end

      # Process each part
      r = / MERGEFIELD "?([^ ]+?)"? (| \\\* MERGEFORMAT )/i
      @parts.each_value do |part|
        part.root.remove_attribute('Ignorable')

        part.xpath('.//w:fldSimple/..').each do |parent|
          parent.children.each do |child|
            next if child.node_name != 'fldSimple'
            instr = child.attribute('instr')

            m = r.match(instr)
            raise Exception, "Could not determine the name of merge field in value #{instr}" if not m

            new_tag = Nokogiri::XML::Node.new("MergeField", part)
            new_tag.content = m[1]
            child.replace(new_tag) # Replace the original convoluted tag with a simplified tag for easy searching a processing
          end
        end

        part.xpath('.//w:instrText/../..').each do |parent|
          begin_tags = parent.xpath('w:r/w:fldChar[@w:fldCharType="begin"]/..')
          end_tags = parent.xpath('w:r/w:fldChar[@w:fldCharType="end"]/..')
          instr_tags = parent.xpath('w:r/w:instrText').map { |x| x.content }

          instr_tags.take(begin_tags.length).each_with_index do |instr, idx|
            m = r.match(instr)
            next if not m

            children = parent.children
            start_idx = children.index(begin_tags[idx]) + 1
            end_idx = children.index(end_tags[idx])
            children[start_idx..end_idx].each do |child|
              child.remove
            end

            new_tag = Nokogiri::XML::Node.new("MergeField", part)
            new_tag.content = m[1]
            begin_tags[idx].replace(new_tag)
          end
        end
      end
    end

    def fields(parts: nil)
      parts = @parts.values if not parts
      fields = Set.new
      parts.each do |part|
        part.xpath('.//w:MergeField').each do |mf|
          fields.add(mf.content)
        end
      end
      fields.to_a
    end

    def merge(parts: nil, **replacements)
      parts = @parts.values if not parts
      parts.each do |part|
        replacements.each do |field, text|
          merge_field(part, field, text)
        end
      end
    end

    def generate()
      clean
      buf = Zip::OutputStream.write_buffer do |out|
        @zip.entries.each do |e|
          unless e.ftype == :directory
            out.put_next_entry(e.name)
            if @parts.keys.include?(e.name)
              out.write @parts[e.name].to_xml(indent: 0).gsub('\n', '')
            else
              out.write e.get_input_stream.read
            end
          end
        end
      end
      buf.seek(0)
      buf
    end

    def write(filename)
      File.open(filename, "w") { |f| f.write(generate().string) }
    end

    def clean(default: '')
      # Remove remaining merge fields
      remain = {}
      fields.each do |field|
        remain[field.to_sym] = default
      end
      merge(**remain)
    end

    private
      def zip_read_xml(filename)
        Nokogiri::XML(@zip.read(filename))
      end

      def merge_field(part, field, text)
        part.xpath(".//w:MergeField[text()=\"#{field}\"]").each do |mf|
          r_elem = Nokogiri::XML::Node.new("r", part)
          t_elem = Nokogiri::XML::Node.new("t", part)
          t_elem.content = text
          t_elem.parent = r_elem
          mf.replace r_elem
        end
      end
  end
end
