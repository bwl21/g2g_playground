infile   = ARGV[0]
basename = File.basename(infile, ".*")


require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'pry'
require 'yaml'

def wrap(s, width = 78)
  paragraphs = s.split("\n\n").map { |ps| ps.gsub(/(.{1,#{width}})(\s+|\Z)/, "\\1\n") }
  paragraphs.join("\n")
end

# encoding: UTF-8

def convert_xlsx_to_genealog_model(workbook)
  labels = workbook[0][1].cells.map { |i| i.value }

  rows  = workbook[0].count - 1
  model = {}
  (2 .. rows).each do |rowid|
    rowmodel = {}
    labels.each_with_index do |key, colid|
      rowmodel[key] = workbook[0][rowid][colid]&.value&.to_s&.strip rescue binding.pry
    end
    model[rowmodel["Ord-Nr"]] = rowmodel unless rowmodel["Ord-Nr"].nil?

  end
  model
end

# jede person hat potenziell eine Familie
# x.y.01  - family_id x.y.01
# x.y.02  - family_id x.y.02
# x.y.1   - family_id x.y.0{ausEhe}
# a.b.c   - family_id a.b.00 - auch hier ist ausEhe 0
# im grunde kann jeder Eintrag zu einer Familie führen
#
# family = {husb: [], wife: [], child: []}   families with members
# person = {husb: [], wife: [], child: []}   role in families

def get_families(model)
  result = model.keys.inject({}) do |result, personid|
    ["", ".00"].each do |i|
      family_id         = "#{personid}#{i}"
      result[family_id] = {id: family_id, child: [], husb: [], wife: []}
    end
    result["..00"] = {id: "..00", child: [], husb: [], wife: []}
    result
  end


  model.keys.inject(result) do |result, personid|
    person  = model[personid]
    famroot = person["Ang-Von"]

   # return(result) if famroot.nil?   # if there is no relative ...

    relatives = model.keys.select { |i| i.match(/^#{personid.gsub(".", "\.")}\.\d+$/) }
    spouses   = relatives.select { |i| i.split(".").last.start_with? "0" }

    famc = nil
    fams = nil
    if person["Art"] == "Ehe" # ehepartner begründen die Familie
      fams = [personid]
    else
      famc = [%Q{#{famroot}.0#{person["Aus-Ehe"]}}] # elternfamilie unless famroot.nil?
      fams = spouses.map { |i| %Q{#{i}} }
    end

    family_id = %Q{#{personid}}
    if person["Art"] == "Ehe"
      if person["Geschl"] == "M"
        result[family_id][:husb] = [personid]
        result[family_id][:wife] = [famroot] unless famroot
      else
        result[family_id][:wife] = [personid]
        result[family_id][:husb] = [famroot] unless famroot
      end
      result[family_id][:date] = person["Hochz-am"]
      result[family_id][:plac] = person["Hochz-in"]
    else # es ist ein Kind
      family_id = %Q{#{famroot}.0#{person["Aus-Ehe"]}}  # Familienzuodnung über Aus-Ehe errechnen
      result[family_id][:child].push(personid) rescue  puts ("familie #{family_id} zu #{personid} fehlt")

      fampart = person["Geschl"] == "M" ? :husb : :wife
      fams.each do |spouse|
        family_id                  = "#{spouse}"
        result[family_id][fampart] = [personid]
      end
    end

    result
  end
  result.delete_if { |k, v| v[:husb].empty? && v[:wife].empty? && v[:child].empty? }
  result
end

def get_family_roles(families)
  personroles = {husb: 'FAMS', wife: 'FAMS', child: 'FAMC'}
  result      = families.inject({}) do |result, (family_id, family)|
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
  id             = personmodel["Ord-Nr"]
  npfx           = personmodel["Titel"]
  name           = personmodel["Name"]
  name           = personmodel["Geb-Name"]
  rufname        = personmodel["Rufname"]&.strip
  vorname        = personmodel["Vorname"]
  vornamerufname = vorname #.gsub(rufname, %Q{*#{rufname.strip}*}) unless rufname.nil?

  sex    = personmodel["Geschl"] == "M" ? "M" : "F"
  ausehe = personmodel["AusEhe"]
  art    = personmodel["AusEhe"]

  beruf = wrap(personmodel["Beruf"] || "", 80).strip

  family_roles = (personmodel[:family_roles] || []).map { |v| %Q{1 #{v[1].to_s.upcase} @#{ $idmapper.fam(v[0])  }@} }.join("\n")
  maiden       = (personmodel[:family_roles] || []).map { |v| %Q{1 #{v[1].to_s.upcase} @#{ $idmapper.fam(v[0])  }@} }.join("\n")

  %Q{
0 @#{$idmapper.indi(id)}@ INDI
1 NAME #{vornamerufname} (#{id}) /#{name}/
2 GiVN #{vornamerufname}
2 SURN #{name}
2 NPFX #{npfx}
2 _RUFNAME #{rufname}
1 NOTE #{beruf.gsub("\n", "\n2 CONT ")}
1 SEX #{sex}
1 REFN #{id}
1 BIRT
2 DATE #{personmodel["Geb-am"]}
2 PLAC #{personmodel["Geb-in"]}
1 DEAT
2 DATE #{personmodel["Gest-am"]}
2 PLAC #{personmodel["Gest-in"]}
  #{family_roles}
  }.strip

end

def get_md_name(personmodel)
  npfx           = personmodel["Titel"]
  name           = personmodel["Name"]
  name           = personmodel["Geb-Name"]
  rufname        = personmodel["Rufname"]&.strip
  vorname        = personmodel["Vorname"]
  id             = personmodel["Ord-Nr"]
  vornamerufname = vorname.gsub(rufname, %Q{*#{rufname.strip}*}) unless rufname.nil?
  %Q{#{name}, #{vornamerufname} (#{id})}
end

def get_md_beruf(personmodel, quote)
  %Q{#{personmodel["Beruf"]&.split("\n")&.join("\n#{quote}")}}
end

def mk_md_death(personmodel)
  if personmodel["Gest-am"]
    gest_in = personmodel["Gst-in"]
    gest_in = gest_in ? " in #{gest_in}": ""
    %Q{- gest. #{personmodel["Gest-am"]} #{gest_in}}
  else
    ''
  end
end

def md_indi(personmodel)
  name      = get_md_name(personmodel)
  md_person = %Q{
<!--  -->
# #{name}

>
>#{get_md_beruf(personmodel, '>')}
>
> - geb. #{personmodel["Geb-am"]} in #{personmodel["Geb-in"]}
> #{mk_md_death(personmodel)}}

  md_hochzeit = personmodel[:relatives]&.select { |i| i["Art"] == 'Ehe' }&.map do |i|
    %Q{
> - Hochzeit am #{i["Hochz-am"]} in #{i["Hochz-in"]} mit
>
>> #{get_md_name(i)}
>>
>>>#{get_md_beruf(i, '>>>')}
>>
>> - geb. #{i["Geb-am"]} in #{i["Geb-in"]}
>> #{mk_md_death(i)}
>>
}
  end

  md_kinder = personmodel[:relatives]&.select { |i| i["Art"] == 'Kind' }&.map do |i|
    %Q{
>> 1. #{get_md_name(i)}
>>
>>    #{get_md_beruf(i, '>>   ')}
>>
>>    - geb. #{i["Geb-am"]} in #{i["Geb-in"]}
>>    #{mk_md_death(i)}
>>}
  end

  unless md_kinder.empty?
    md_kinder = %Q{
> -  **Kinder**

#{md_kinder.join}
    }
  end

  [md_person, md_hochzeit, md_kinder].join("")
end

# expose a particluar family as gedcom
def ged_fam(family_id, family)
  child = family[:child].map { |c| %Q{1 CHIL @#{$idmapper.indi(c)}@} }
  husb  = family[:husb].map { |c| %Q{1 HUSB @#{$idmapper.indi(c)}@} }
  wife  = family[:wife].map { |c| %Q{1 WIFE @#{$idmapper.indi(c)}@} }

  %Q{
0 @#{$idmapper.fam(family_id)}@ FAM
#{[child, husb, wife].flatten.compact.join("\n")}}.strip
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
      result = id.gsub(/[^a-zA-Z0-9_]/, "_")
      #@nextnumber += 1
      @idmap[id] = result
    end
    %Q{#{clazz}#{result}}
  end
end


def patch_relatives(model)
  relatives = model.group_by { |key, person| person["Ang-Von"] }
  # .last comes from the somehow complex model
  model.keys.each { |i| model[i][:relatives] = relatives[i]&.map { |j| j.last } }
end


#####################################################

$idmapper = IdSanitizer.new


if File.extname(infile) == ".json"
  model = JSON.load(File.read(infile))
else
  workbook = RubyXL::Parser.parse(infile)
  model = convert_xlsx_to_genealog_model(workbook)
  File.open("inputs/#{basename}.json", "w:UTF-8") do |f|
    f.puts(JSON.pretty_generate model)
  end
end


families     = get_families(model)
family_roles = get_family_roles(families)

patch_relatives(model)
patch_family_roles(family_roles, model)
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

  f.puts "0 TRLR" # was requested by yed
end


File.open("mdreport/#{basename}.test.md", "w:UTF-8") do |f|
  model.keys.each do |id|
    f.puts md_indi(model[id]) if model[id][:relatives]
  end
end
