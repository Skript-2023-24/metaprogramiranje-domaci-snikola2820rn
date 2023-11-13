require "google_drive"
require "debug"

class GoogleSheet
  attr_accessor :ws, :worksheets

  def initialize
    @session = GoogleDrive::Session.from_config("config.json")
    @spreadsheet = @session.spreadsheet_by_key("1euDwwCDhUA6VWmFpeWurToRwUMduraOY1nG_ZpAC4yA")
    @worksheets = []
    @spreadsheet.worksheets.each do |ws|
      @worksheets << Worksheet.new(ws,self) if ws.num_rows != 0
    end
    @ws = @worksheets[0]
  end

  class Worksheet
    attr_accessor :ws, :start, :empty, :cols, :rows, :gs

    # t[prvaKolona] mora da vrati neki objekat
    def initialize(ws,gs)
      @gs=gs
      @ws = ws
      @start = {}
      rows = ws.rows
      tmp = rows.index { |row| !row.all?{ |element| element.nil? || element == '' }}
      @start[:row] = (tmp ? tmp : 0)
      @cols = {}
      rows[@start[:row]].each_with_index do |element, i|
        if element!=nil and element!=''
          @cols[element.to_sym] = i+1
        end
      end
      #prvi red, 0-ti element u nizu, redovi krecu od 1
      @rows = []

      for i in (@start[:row]+1)...rows.length do
        all =  rows[i].all? { |element| element.nil? || element == ''}
        any = rows[i].any? { |element| element.downcase.include? "total" }
        if !all and !any
          @rows << (i+1)
        end
      end
      #pretpostavljam da tabela ne moze da
      #se menja van programa dok radimo sa njom u programu
    end

    def row(row)
      content = []
      @cols.each do |key, value|
        content << @ws[@rows[row-1],value]
      end
      content
    end

    def [](col)
      Column.new(self,col)
    end

    def []=(row,col,value)
      # if row >= @rows.length
      #   @rows << @rows[-1] + 1 + row - @rows.length
      # end
      @ws[@rows[row],col]=value
      @ws.save
    end

    class Column
    include Enumerable
      def initialize (ws,col)
        begin
          if col.instance_of? Symbol
            @col = ws.cols.transform_keys do |key|
              key.to_s.gsub(/\s+/, "").downcase.to_sym
            end [col.to_s.downcase.to_sym]
          else
            @col = ws.cols[col.to_sym]
          end
        rescue KeyError
          puts "Ne postoji kolona #{col}"
          return
        end
        @content = []
        @ws = ws
        ws.rows.each do |row|
          @content << ws.ws[row,@col]
        end
      end

      def content=(value)
        @content = value
        @ws.rows.each_with_index do |row, i|
          ws[row,@col] = value[i]
        end
      end

      def [] (row)
        @content[row]
      end

      def []= (row, value)
        @content[row] = value
        @ws[row,@col] = value
      end

      def each (&block)
        @content.each(&block)
      end

      def to_s
        @content.to_s
      end

      def sum
        # radice i ako u koloni ima stringova
        @content.inject(0){|sum,x| sum + x.to_i }
      end

      def avg
        self.sum / @content.length.to_f
      end
      # def map (&block)
      #   result = super &block
      #   @content = result
      # end
      def method_missing(key, *args)
        text = key.to_s
        @ws.row(@content.map(&:downcase).index(text)+1)
      end
    end

    def each
      @rows.each do |row|
        @cols.each do |key,value|
          yield @ws[row,value]
        end
      end
    end

    def all
      content = []
      line = []
      @rows.each do |row|
        @cols.each do |key,value|
          line << @ws[row,value]
        end
        content << line
        line = []
      end
      content
    end

    def + (other)
      if @cols.keys.sort != other.cols.keys.sort
        puts "Kolone nisu iste"
        return
      end

      new_ws = @ws.spreadsheet.add_worksheet("Sheet " + (@ws.spreadsheet.worksheets.length+1).to_s,
        @ws.num_rows > other.ws.num_rows ? @ws.num_rows : other.ws.num_rows,
        @ws.num_cols)

      new_ws.insert_rows(1, [@cols.keys])
      @ws.save

      @rows.each do |row|
        new_ws.list.push(@cols.transform_keys(&:to_s).transform_values { |v| @ws[row,v] })
      end

      other.rows.each do |row|
        new_ws.list.push(other.cols.transform_keys(&:to_s).transform_values { |v| other.ws[row,v] })
      end
      new_ws.save()
      @gs.add_ws(new_ws)
    end

    def - (other)

      if @cols.keys.sort != other.cols.keys.sort
        puts "Kolone nisu iste"
        return
      end
      new_ws = @ws.spreadsheet.add_worksheet("Sheet " + (@ws.spreadsheet.worksheets.length+1).to_s, @ws.num_rows, @ws.num_cols)
      new_ws.insert_rows(1, [@cols.keys])
      @ws.save

      @rows.each do |row1|
        cols1 = @cols.transform_values { |v| @ws[row1,v] }
        flag = true
        other.rows.each do |row2|
          cols2 = other.cols.transform_values { |v| other.ws[row2,v] }
          if cols1 == cols2
            flag = false
            break
          end
        end
        new_ws.list.push(cols1.transform_keys(&:to_s)) if flag
      end
      new_ws.save
      @gs.add_ws(new_ws)
    end

    def method_missing(key, *args)
      self[key]
    end

  end

  def add_ws (ws)
    @worksheets<<Worksheet.new(ws,self)
  end
end



sheet = GoogleSheet.new

# p sheet.ws.all
# p sheet.ws.row(1)

# sheet.ws.each { |element| puts element }

puts sheet.ws["A"][2]
puts sheet.ws["B"][1]
sheet.ws["B"][1] = "test"

# puts sheet.ws.b
# puts sheet.ws.b.sum
# puts sheet.ws.b.avg
# puts sheet.ws.b.bar
# p sheet.ws.b.map { |element| element.to_i + 1 }
# p sheet.ws.b.reduce(0) { |sum, element| sum + element.to_i }


# sheet.worksheets[0] - sheet.worksheets[1]
# sheet.worksheets[0] + sheet.worksheets[1]
