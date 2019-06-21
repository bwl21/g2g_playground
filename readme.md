# playground for a gedcom converter

Pupose of this project is to share experiments with GEDCOM. Beyond this,
it is of no practical use.

# getting started

## setup

* clone this project
* using ruby 2.5 install the requested gems
 
```ruby
cd ruby
bundle install
``` 

## create the GEDCOM file

```bash

ruby ruby/g2g.rb inputs/gedcomtest.xslx
```

inspect the resulting file `gedcom/gedcomtest.ged`

it also craets `inputs/gedcomtest.debug.yaml` for debugging purpoeses


# todo

let it work with yaml input as well



