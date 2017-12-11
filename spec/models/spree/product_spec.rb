require 'spec_helper'

RSpec.describe Spree::Product do
  describe "quiero que funcione" do
      let(:product){ Spree::Product.new }
      it "imprime" do
        puts product.inspect
      end

  end
end
