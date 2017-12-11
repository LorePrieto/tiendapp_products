require "bundler/setup"
require "tiendapp_products"
require 'factory_bot'

ENV['RAILS_ENV'] ||= 'test'

require File.expand_path('../dummy/config/environment', __FILE__)

require 'rspec/rails'

require 'spree/testing_support/factories'
require 'spree/testing_support/preferences'
require 'spree/testing_support/controller_requests'
require 'spree/testing_support/authorization_helpers'
require 'spree/testing_support/url_helpers'

RSpec.configure do |config|
  config.include FactoryBot::Syntax::Methods
  config.include Spree::TestingSupport::ControllerRequests, type: :controller
  config.include Spree::TestingSupport::Preferences
  config.include Spree::TestingSupport::UrlHelpers
end
