# Class to create custom metadata for 'Word Count' and 'Words Per Page.'
# @author mrk
# @version 1.0
# * +@opts+ contains the options
# * +@dialog+ is an Nx::ProcessDialog
# * +@metadata+ is a hash of { field => {value => items} }
class WordCounter
  # Identifies PDFs with text, and adds custom metadata for:
  #  Word Count
  #  Words Per Page
  #
  # @param opts [Hash] the options
  # @option opts [String] :query the search query
  # @option opts [Boolean] :handleExcluded if only items with no exclusion
  # @option opts [Boolean] :recalculate if recalculating values
  def initialize(opts)
    Nx::ProgressDialog.forBlock do |progress_dialog|
      @opts = opts
      @dialog = progress_dialog
      @dialog.setTitle('Word Count Custom Metadata')
      @dialog.setAbortButtonVisible(false)
      @dialog.setLogVisible(true)
      @dialog.setTimestampLoggedMessages(true)
      @metadata = fields_hash(['Word Count', 'Words Per Page'])
      run
    end
  end

  # Adds Word Count and Words Per Page custom metadata.
  def annotate
    @dialog.setMainStatus('Annotating Items')
    @metadata.each_with_index do |(n, h), i|
      @dialog.setMainProgress(2 + i)
      put_metadata(n, h)
      break if @dialog.abortWasRequested
    end
  end

  # Completes the dialog, or logs the abortion.
  def close
    if @dialog.abortWasRequested
      @dialog.logMessage('Aborted')
    else
      @dialog.setCompleted
    end
  end

  # Sorts the item by words and words per page.
  # Word count determined by String.gsub(/[^-a-zA-Z]/, ' ').split.size.
  #
  # @param item [Item] a Nuix item
  def count(item)
    c = item.get_text_object.to_string.gsub(/[^-a-zA-Z]/, ' ').split.size
    @metadata['Word Count'][c] << item
    avg = c / pages(item).to_f
    @metadata['Words Per Page'][avg] << item
  end

  # The hash of custom metadata fields and values.
  #
  # @param names [Array] the custom metadata field names
  # @return [Hash] of { field => {value => items} }
  def fields_hash(names)
    Hash[names.collect { |n| [n, Hash.new { |h, k| h[k] = [] }] }]
  end

  # Finds non-searchable PDFs.
  #
  # @return [Set<Item>] results of the search
  def find_items
    @dialog.setMainStatus('Finding Items')
    @dialog.setMainProgress(0, 4)
    q = query
    @dialog.setSubStatusAndLogIt("Searching for: #{q}")
    @dialog.setSubProgress(0, 1)
    $current_case.search_unsorted(q)
  end

  # The page count of an item.
  #
  # @param item [Item] a PDF item
  # @return [Integer] page count
  def pages(item)
    count = item.get_properties.get('Page Count')
    return count unless count.nil?

    1
  end

  # Adds custom metadata.
  #
  # @param name [String] custom metadata field name
  # @param hash [Hash] hash of {value => items}.
  def put_metadata(name, hash)
    @dialog.setSubStatusAndLogIt("Adding \"#{name}\" Custom Metadata")
    @dialog.setSubProgress(0, hash.size)
    ba = $utilities.get_bulk_annotater
    hash.each_with_index do |(v, i), index|
      break if @dialog.abortWasRequested

      @dialog.setSubProgress(index)
      ba.put_custom_metadata(name, v, i, nil)
    end
  end

  # The search query to identify items to process.
  # Includes 'has-exclusion:0' if @opts[:handleExcluded].
  # Includes '-custom-metadata-exact:...' if @opts[:recalculate].
  #
  # @return [String] search query string
  def query
    terms = ["(#{@opts[:query]})"]
    terms << 'has-exclusion:0' if @opts[:handleExcluded]
    unless @opts[:recalculate]
      @metadata.each_key { |k| terms << "NOT custom-metadata-exact:\"#{k}\":*" }
    end
    terms.join(' AND ')
  end

  # Finds items, calculates fields, then annotates.
  def run
    items = find_items
    if items.empty?
      @dialog.logMessage('No items found')
    else
      sort(items)
      annotate unless @dialog.abortWasRequested
    end
    close
  end

  # Counts the words and sorts the items.
  #
  # @param items [Set<Items>] the items to check
  def sort(items)
    sort_dialog(items.size)
    items.each_with_index do |i, index|
      break if @dialog.abortWasRequested

      @dialog.setSubProgress(index)
      count(i)
    end
  end

  # Initializes @dialog for sort.
  #
  # @param size [Integer] the number of items
  def sort_dialog(size)
    @dialog.setMainStatus('Counting Words')
    @dialog.setMainProgress(1)
    @dialog.setSubStatusAndLogIt("Sorting #{size} items")
    @dialog.setSubProgress(0, size)
    @dialog.setAbortButtonVisible(true)
  end

  # Nx Module
  module Nx
    require File.join(__dir__, 'Nx.jar')
    java_import 'com.nuix.nx.NuixConnection'
    java_import 'com.nuix.nx.LookAndFeelHelper'
    java_import 'com.nuix.nx.dialogs.ProgressDialog'
    LookAndFeelHelper.setWindowsIfMetal
    NuixConnection.setUtilities($utilities)
    NuixConnection.setCurrentNuixVersion(NUIX_VERSION)
  end
end
