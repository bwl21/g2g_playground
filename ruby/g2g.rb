
infile = ARGV[0]
basename = File.basename(infile, ".xlsx")


require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'pry'
require 'yaml'

# encoding: UTF-8

def convert_xlsx_to_genealog_model(workbook)
  labels = workbook[0][1].cells.map { |i| i.value }

  rows  = workbook[0].count - 1
  model = {}
  (2 .. rows).each do |rowid|
    rowmodel = {}
    labels.each_with_index do |key, colid|
      rowmodel[key] = workbook[0][rowid][colid]&.value rescue binding.pry
    end
    model[rowmodel["Ord-Nr"]] = rowmodel unless rowmodel["Ord-Nr"].nil?

  end
  model
end

# jede person hat potenziell eine Familie
# x.y.01  - fam x.y.01
# x.y.02  - fam x.y.02
# x.y.1   - fam x.y.0{ausEhe}
# a.b.c   - fam a.b.00 - auch hier ist ausEhe 0
# im grunde kann jeder Eintrag zu einer Familie führen
#
# family = {husb: [], wife: [], child: []}   families with members
# person = {husb: [], wife: [], child: []}   role in families

def get_families(model)
  result = model.keys.inject({}) do |result, personid|
    ["", ".00"].each do |i|
      fam         = "#{personid}#{i}"
      result[fam] = {id: fam, child: [], husb: [], wife: []}
    end
    result["..00"] = {id: "..00", child: [], husb: [], wife: []}
    result
  end


  model.keys.inject(result) do |result, personid|
    person  = model[personid]
    famroot = person["Ang-Von"]

    relatives = model.keys.select { |i| i.match(/^#{personid.gsub(".", "\.")}\.\d+$/) }
    spouses   = relatives.select { |i| i.split(".").last.start_with? "0" }

    if person["Art"] == "Ehe" # ehepartner begründen die Familie
      fams = [personid]
    else
      famc = [%Q{#{famroot}.0#{person["Aus-Ehe"]}}] # elternfamilie
      fams = spouses.map { |i| %Q{#{i}} }
    end


    if person["Art"] == "Ehe"
      if person["Geschl"] == "M"
        result[personid][:husb] = [personid]
        result[personid][:wife] = [famroot]
      else
        result[personid][:wife] = [personid]
        result[personid][:husb] = [famroot]
      end
    else # es ist ein Kind
      fam = %Q{#{famroot}.0#{person["Aus-Ehe"]}}
      result[fam][:child].push(personid) rescue puts ("familie #{fam} zu #{personid} fehlt")

      fampart = person["Geschl"] == "M" ? :husb : :wife
      fams.each do |spouse|
        result[spouse][fampart] = [personid]
      end
    end

    result
  end
  result.delete_if { |k, v| v[:husb].empty? && v[:wife].empty? && v[:child].empty? }
  result
end

def get_family_roles(families)
  personroles={husb: 'FAMS', wife: 'FAMS', child: 'FAMC'}
  result = families.inject({}) do |result, (family_id, family)|
    [:husb, :wife, :child].each do |role|
      family[role].each do |person_id|
        result[person_id] ||= []
        result[person_id].push([family_id, personroles[role]])
      end
    end
    result
  end
  result
end

def patch_family_roles(family_roles, model)
   family_roles.each do |k, v|
     model[k][:family_roles] = v
   end
end

# expose an individual as gedcom
def get_indi(personmodel)

  id      = personmodel["Ord-Nr"]
  name    = personmodel["Name"]
  vorname = personmodel["Vorname"]
  sex     = {"M" => "M", "W" => "F"}[personmodel["Geschl"]]
  ausehe  = personmodel["AusEhe"]
  art     = personmodel["AusEhe"]


  family_roles = (personmodel[:family_roles] || []).map{|v| %Q{1 #{v[1].to_s.upcase} @#{$idmapper.fam(v[0])}@}}.join("\n")

  %Q{
0 @#{$idmapper.indi(id)}@ INDI
1 NAME #{vorname} /#{name}/ (#{id})
1 NOTE #{personmodel["Beruf"]}
1 SEX #{sex}
1 BIRT
2 DATE #{personmodel["Geb-am"]}
2 PLAC #{personmodel["Geb-in"]}
1 DEAT
2 DATE #{personmodel["Gest-am"]}
2 PLAC #{personmodel["Gest-in"]}
#{family_roles}
  }.strip

end


# expose a particluar family as gedcom
def ged_fam(family_id, family)
  child = family[:child].map { |c| %Q{1 CHIL @#{$idmapper.indi(c)}@} }
  husb  = family[:husb].map { |c| %Q{1 HUSB @#{$idmapper.indi(c)}@} }
  wife  = family[:wife].map { |c| %Q{1 WIFE @#{$idmapper.indi(c)}@} }
  %Q{
0 @#{$idmapper.fam(family_id)}@ FAM
#{[child, husb, wife].flatten.compact.join("\n")}}

end


def ged_header
  %Q{0 HEAD
1 CHAR UTF-8
1 COPR Bernhard Weichel
1 DATE 13.6.2019
2 TIME 14:19:43
1 DEST ANSTFILE
1 FILE Muster_GEDCOM.ged
1 GEDC
2 FORM LINEAGE-LINKED
2 VERS 5.5.1
1 LANG German
}
end

class IdSanitizer

  def initialize
    @idmap      = {}
    @nextnumber = 10000
  end

  def nextid

  end

  def indi(id)
    to_gedcom(id, "I")
  end

  def fam(id)
    to_gedcom(id, "F")
  end

  def to_gedcom(id, clazz = "X")
    result = @idmap[id]
    unless result
      result     = id.gsub(/[^a-zA-Z0-9_]/, "_")
      #@nextnumber += 1
      @idmap[id] = result
    end
    %Q{#{clazz}#{result}}
  end
end


$idmapper = IdSanitizer.new

workbook = RubyXL::Parser.parse(infile)

model = convert_xlsx_to_genealog_model(workbook)

families     = get_families(model)
family_roles = get_family_roles(families)

File.open("inputs/#{basename}.debug.yaml", "w:UTF-8") do |f|
  f.puts({families: families, family_roles: family_roles, model: model}.to_yaml)
end



File.open("gedcom/#{basename}.ged", "w:UTF-8") do |f|

  f.puts ged_header

  model.each do |k, v|
    f.puts get_indi(v)
  end

  families.each do |k, v|
    f.puts ged_fam(k, v)
  end

  f.puts "0 TRLR"   # was requested by yed

end
