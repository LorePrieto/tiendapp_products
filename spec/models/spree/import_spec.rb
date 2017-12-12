require 'spec_helper'
require 'roo'

RSpec.describe TiendappProducts::Import do
  describe "From an excel the correct database enttries are created with Spree" do
    it "algo" do
      puts TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')
    end
  end
end
