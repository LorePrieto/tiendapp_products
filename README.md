# TiendappProducts

This gem is created for importing and export product using Spree. Its functionailies where thought specifically for the TienApp's app.

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'tiendapp_products'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install tiendapp_products

## Usage

For importing do
```ruby
TiendappProducts::Import.create_products([path_to_excel])
```
Note: the excel needs a specific format. (Look under 'spec/fixtures/imported.xlsx')
For exporting do
```ruby
TiendappProducts::Export.get_products([path_to_save])
```
Note: the saved file needs to be a .xlsx, so the path has to be like 'spec/fixtures/exported.xlsx'.
Note 2: no validation of correct data is done when importing.

## Development

After checking out the repo, run `bin/setup` to install dependencies. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

For testing do

    $ bundle install
    $ bundle exec rake test_app
    $ bundle exec rspec spec

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/tiendapp_products.
