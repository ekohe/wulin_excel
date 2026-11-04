ActiveRecord::Migration.verbose = false
ActiveRecord::Schema.define do
  create_table :countries do |t|
    t.string :name
  end
  create_table :people do |t|
    t.string :first_name
    t.integer :age
    t.datetime :birthdate
    t.boolean :vip, default: false
    t.text :signature
    t.references :country
  end
  create_table :posts do |t|
    t.string :name
    t.references :person
  end
end

ActiveRecord::Base.connection.execute("CREATE VIEW person_summaries AS SELECT id, first_name, country_id FROM people")

class Country < ActiveRecord::Base
end

class Person < ActiveRecord::Base
  belongs_to :country, optional: true
  has_many :posts
end

class Post < ActiveRecord::Base
  belongs_to :person, optional: true
end

# A database view has no primary key, like bss ObserverSeatCommentaryPosition.
class PersonSummary < ActiveRecord::Base
  self.primary_key = nil
  belongs_to :country, optional: true
end

class ApplicationController < ActionController::Base
  def self.define_menu; end
end

class PersonGrid < WulinMaster::Grid
  title "People"
  model Person
  default_sorting_state({column: "people.first_name", direction: "ASC"})

  column :first_name
  column :age
  column :birthdate
  column :vip
  column :signature
  column :country
end

class PersonScreen < WulinMaster::Screen
  title "People"
  path "/people"
  grid PersonGrid
end

class PeopleController < WulinMaster::ScreenController
  controller_for_screen PersonScreen
end

# A grid that adds its own multi-column SELECT with an aggregate, as ceres and atlas screens do.
class AggregateGrid < WulinMaster::Grid
  title "Aggregate people"
  model Person
  default_sorting_state({column: "people.first_name", direction: "ASC"})

  column :first_name
  column :posts_count, sql_expression: "(SELECT COUNT(*) FROM posts WHERE posts.person_id = people.id)"
end

class AggregateScreen < WulinMaster::Screen
  title "Aggregate people"
  path "/aggregates"
  grid AggregateGrid
end

class AggregatesController < WulinMaster::ScreenController
  controller_for_screen AggregateScreen
  add_callback :query_initialized, :add_posts_count

  private

  def add_posts_count
    @query = Person.select("people.*, (SELECT COUNT(*) FROM posts WHERE posts.person_id = people.id) AS posts_count")
  end
end

# Same screen with a permission rule that always denies.
class ForbiddenScreen < WulinMaster::Screen
  title "Forbidden people"
  path "/forbiddens"
  grid PersonGrid

  def authorized?(_user = nil)
    false
  end
end

class ForbiddensController < WulinMaster::ScreenController
  controller_for_screen ForbiddenScreen

  def current_user
    Object.new
  end
end

class PersonSummaryGrid < WulinMaster::Grid
  title "Person summaries"
  model PersonSummary
  default_sorting_state({column: "person_summaries.first_name", direction: "ASC"})

  column :first_name
  column :country
end

class PersonSummaryScreen < WulinMaster::Screen
  title "Person summaries"
  path "/person_summaries"
  grid PersonSummaryGrid
end

class PersonSummariesController < WulinMaster::ScreenController
  controller_for_screen PersonSummaryScreen
end

# A callback that swaps the query for another model, like bss SadWaitingsController.
class UnfilteredSummaryScreen < WulinMaster::Screen
  title "Unfiltered summaries"
  path "/unfiltered_summaries"
  grid PersonSummaryGrid
end

class UnfilteredSummariesController < WulinMaster::ScreenController
  controller_for_screen UnfilteredSummaryScreen
  add_callback :query_ready, :return_empty_result_set

  private

  def return_empty_result_set
    @query = Person.where("false")
  end
end

Rails.application.routes.draw do
  resources :unfiltered_summaries, only: :index
  resources :person_summaries, only: :index
  resources :people, only: :index
  resources :aggregates, only: :index
  resources :forbiddens, only: :index
end
