(asdf:defsystem #:xlsx
  :name "XLSX"
  :description "Basic reader for Excel files"
  :author "Carlos Ungil <ungil@mac.com>"
  :license "MIT"
  :serial t
  :components ((:file "package")
               (:file "xlsx"))
  :depends-on (:zip :flexi-streams :xmls))

