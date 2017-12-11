source "https://rubygems.org"

git_source(:github) {|repo_name| "https://github.com/#{repo_name}" }

group :development, :test do
  spree_branch = '3-0-stable'
  gem 'spree',                  github: 'spree/spree',                  branch: spree_branch
  gem 'sqlite3'
end

# Specify your gem's dependencies in tiendapp_products.gemspec
gemspec
