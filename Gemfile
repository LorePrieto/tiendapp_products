source "https://rubygems.org"

git_source(:github) {|repo_name| "https://github.com/#{repo_name}" }

group :development, :test do
  gem 'spree', '~> 3.0.4'
  gem 'spree_i18n', github: 'spree-contrib/spree_i18n', branch: '3-0-stable'
  gem 'sqlite3'
end

# Specify your gem's dependencies in tiendapp_products.gemspec
gemspec
