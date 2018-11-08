# Menu Title: PDF Word Counter
# Needs Case: true
opts = {
  query: 'mime-type:application/pdf AND content:*',
  handleExcluded: true,
  recalculate: false
}
begin
  require File.join(__dir__, 'word_counter.rb')
  WordCounter.new(opts)
end
