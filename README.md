# DocxMailmerge

Performs a mail merge on DOCX (Microsoft Office Word) files. The user can use the Mail Merge Manager feature of Microsoft Word to create templates with merge fields embedded in the document. This library then allows a developer to take that DOXC file template as an input along with a hash containing the merge field names and values, and will return a merged document as a DOCX file that can be saved as a file, or persisted in a paperclip attachment or any other programmatic use.

DocxMailmerge is a Ruby port of the Python [docx-mailmerge](https://github.com/Bouke/docx-mailmerge/) library. It currently only supports simple merge fields (without inner formatting) and does not support image merge.

## Installation

Add this line to your application's Gemfile:

    gem 'docx_mailmerge'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install docx_mailmerge

## Usage

To begin, open a merge file:

    require 'docx_mailmerge'
    document = DocxMailmerge::MailMerge.new('doc.docx')

To figure out which fields are available, list all merge fields:

    puts document.fields

Do the actual merge:

    document.merge(foo: 'Test')

Get the output stream:

    document.generate

Alternatively, write to a file:

    document.write('doc_merged.docx')

## Contributing

1. Fork it ( https://github.com/[my-github-username]/docx_mailmerge/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
