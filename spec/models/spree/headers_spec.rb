require 'spec_helper'
require 'roo'

RSpec.describe TiendappProducts::Export do
  describe "Exported excel have correct headers" do
      it "should create an excel with correct headers" do
        TiendappProducts::Export.get_products()
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        expect(xlsx.sheets).to eql(["Productos", "Propiedades", "Opciones", "Variantes", "Stock"])
        expect(xlsx.sheet("Productos").row(1)).to eql(["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías" ])
        expect(xlsx.sheet("Propiedades").row(1)).to eql(["ID Producto", "Propiedad", "Valor"])
        expect(xlsx.sheet("Opciones").row(1)).to eql(["ID Producto", "Opción", "Valores" ])
        expect(xlsx.sheet("Variantes").row(1)).to eql(["ID", "ID Producto", "Opciones", "SKU", "Precio", "Peso", "Altura", "Longitud", "Profundidad"])
        expect(xlsx.sheet("Stock").row(1)).to eql(["ID Producto", "ID Variante", "Cantidad", "Ubicación", "Backorderable"])
      end
  end
end
