(ns hvitmiddleware.core
(:require 
            [noir.response :as resp]
            [clj-http.client :as client]
            ))

(defn foo
  "I don't do a whole lot."
  [x]
  (println x "Hello, World!"))

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
  "Creates a SQL query using paging and ROWNUM"
  (str "SELECT * from (select " (clojure.string/join "," (map #(str "a." %) properties))
    ", ROWNUM rnum from (select " (clojure.string/join "/" properties)
    " from " table
    " order by " (clojure.string/join "," order) " ) a "
    " WHERE ROWNUM <= " max
    ") WHERE " (if-not predicate "" (str predicate " and ")) " rnum >= " from))  
   
