require 'spec_helper'
require 'roo'

RSpec.describe TiendappProducts::Import do
  describe "From an excel the correct database enttries are created with Spree" do
    it "should create the correct database entries for simple case" do
      #We load the excel
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')

      #Check the database for the correct entries
      product = Spree::Product.where(slug: "queque-en-molde-de-cupcake").first
      sc = Spree::ShippingCategory.where(name: "Por defecto").first
      #Products
      expect(product.name).to eql("queque")
      expect(product.description).to eql("en molde de cupcake")
      expect(product.price).to eql(1500)
      expect(product.meta_description).to eql("Este es el mejor queque de Chile")
      expect(product.available_on.strftime("%Y-%m-%d")).to eql(DateTime.parse("2017-12-06 17:26:02 UTC").strftime("%Y-%m-%d"))
      expect(product.shipping_category_id).to eql(sc.id)
      #Propiedades
      expect(product.properties.first.name).to eql("Hecho en casa")
      expect(product.properties.first.presentation).to eql("Sí")
      expect(product.properties.second.name).to eql("For real no fake")
      expect(product.properties.second.presentation).to eql("No")
      #Options
      ot1 = product.option_types.first
      expect(ot1.name).to eql("Sabor")
      expect(ot1.option_values.first.name).to eql("Nueces")
      expect(ot1.option_values.second.name).to eql("Vainilla")
      ot2 = product.option_types.second
      expect(ot2.name).to eql("Tamaño")
      expect(ot2.option_values.first.name).to eql("estoy a dieta")
      expect(ot2.option_values.second.name).to eql("gigante")
      #variants
      #Locations
      #Stock
    end
  end
end
