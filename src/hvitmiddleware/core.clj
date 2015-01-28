(ns hvitmiddleware.core

(:import 
           (hvitmiddleware.java ExcelHelper)
           
 )
           
(:require 
            [noir.response :as resp]
            [clj-http.client :as client]
            ))

(defn foo
  "I don't do a whole lot."
  [x]
  (println x "Hello, World!"))
  
(defn make-excel [fileName header_arr rowdata sum title headerheight headercols  isall url  extraParams  rowname  pager]
 ;(println "hello world")
 (try
      
        (ExcelHelper/writeExcel fileName header_arr rowdata sum title headerheight headercols  isall url  extraParams  rowname  pager) 
      (catch Exception ex
        {:isok false :msg (.getMessage ex)})
       )
 
 ) 

(defn session-filter [handler url]
  (fn [req]
    (let [
          params  (:params req)
          sessionid (:sessionid params)
          session  (if (nil? sessionid) {:body nil} (client/get (str url "/auth/sessioncheck")
                     {:cookies {"ring-session" {:value sessionid }} :as :clojure}))
          ]
      (if (nil? (:body session)) (resp/json {:success false :msg "未登录，无权限"})(handler req))
      )
    ;(timbre/debug req)
    )
  


  )
  
(defn create-oraclequery-paging [{:keys [table properties order predicate from max] :or {max 100} }]
  "Creates a ORLCESQL query using paging and ROWNUM"
  (str "SELECT * from (select " (clojure.string/join "," (map #(str "a." %) properties))
    ", ROWNUM rnum from (select " (clojure.string/join "," properties)
    " from " table
    (if-not predicate "" (str " where " predicate "  "))
    " order by " (clojure.string/join "," order) " ) a "
    " WHERE ROWNUM <= " max
    ") WHERE "  " rnum >= " from))
   
(defn get-oraclequery-total [{:keys [table predicate ]}]
  "Get a ORLCESQL query totalnum"
  (str "SELECT count(*) "
    " from " table
    " " (if-not predicate "" (str " WHERE " predicate " "))))

