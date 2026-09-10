#!/usr/bin/env nbb
;; scripts/mutate-surface-agreement.cljs
;;
;; Proves that verify-surface-agreement.cljs discriminates. For each pinned fact
;; it copies the tree to a scratch directory, breaks exactly that one thing, and
;; requires that the verifier (a) exits 1 and (b) names the check that was broken.
;;
;;   nbb scripts/mutate-surface-agreement.cljs
;;
;; (b) is the point. A mutation that turns the tree red for some *other* reason
;; -- a reader that now throws, a file that no longer parses, a renamed symbol
;; that happens to contain the string being searched for -- looks exactly like a
;; successful demonstration from the exit code alone. This workspace has made
;; that mistake four times in one day (CLAUDE.md, "壊したものと報告されたものが
;; 一致することを確かめる"). So each case declares the check ids it must turn
;; red, and a case that goes red for anything else is itself a failure here.
;;
;; Writing those declarations is where the couplings got found: three cases fire
;; two checks, and each pair is a real implication rather than noise. They are
;; annotated at the case.
;;
;; Exit codes: 0 every mutation was caught by the intended check · 1 one was not.

(ns mutate-surface-agreement
  (:require ["fs" :as fs]
            ["path" :as path]
            ["os" :as os]
            ["child_process" :as cp]
            [clojure.string :as str]))

(def root (path/resolve (.cwd js/process)))
(def verifier (path/join root "scripts" "verify-surface-agreement.cljs"))

(def O (path/join "appview" "outlook-mcp-component"))
(def G (path/join "appview" "gmail-mcp-component"))
(def gmail-bin (path/join G "gmail-mcp-component"))

;; Each case: break exactly one thing, name the check(s) that must catch it.
(def cases
  [{:id "wrangler main now points at the facade"
    :edit [:sub (path/join O "wrangler.jsonc")
           "svelte/.svelte-kit/cloudflare/_worker.js" "src/app.ts"]
    :must #{:deploy-target :facade-is-not-the-deploy-target}}

   {:id "a SvelteKit server route appears"
    :edit [:write (path/join O "svelte" "src" "routes" "+server.ts")
           "export const GET = () => new Response('ok');\n"]
    :must #{:deployed-worker-has-no-server-routes}}

   ;; The next two move the facade prefix and the UI namespace *towards* each
   ;; other, so :ui-urls-the-facade-would-serve necessarily lifts off 0 too.
   ;; That co-firing is the check doing its job -- it is the one that answers
   ;; "would the facade serve the browser if it were deployed" -- so it is
   ;; declared rather than suppressed.
   {:id "the facade proxies a different namespace"
    :edit [:sub (path/join O "src" "app.ts")
           "com.etzhayyim.apps.outlook." "etzhayyim.outlook.v1."]
    :must #{:facade-nsid-prefix :ui-urls-the-facade-would-serve}}

   {:id "the UI calls a different namespace"
    :edit [:sub (path/join O "svelte" "src" "App.svelte")
           "/xrpc/etzhayyim.outlook.v1.OutlookService" "/xrpc/com.etzhayyim.apps.outlook"]
    :must #{:ui-service-base :ui-urls-the-facade-would-serve}}

   ;; Switching the UI to a slash also lifts UI∩e2e from 0 to 1 (GetConnection
   ;; is the one method name they share), which is the same edit seen from the
   ;; test suite's side.
   {:id "the UI switches to a slash separator"
    :edit [:sub (path/join O "svelte" "src" "App.svelte")
           "${SERVICE_BASE}.${method}" "${SERVICE_BASE}/${method}"]
    :must #{:ui-joins-method-with-dot :ui-urls-the-e2e-suite-exercises}}

   {:id "the UI renames one of its six calls"
    :edit [:sub (path/join O "svelte" "src" "App.svelte")
           "await callApi(\"Disconnect\")" "await callApi(\"Revoke\")"]
    :must #{:ui-methods}}

   {:id "the outlook manifest stops naming the facade"
    :edit [:sub (path/join O "kotodama.jsonld")
           "\"path\": \"src/app.ts\"" "\"path\": \"svelte/.svelte-kit/cloudflare/_worker.js\""]
    :must #{:outlook-manifest-points-at-the-facade}}

   ;; Repointing the gmail manifest at the file that IS committed also settles
   ;; the absence check, because that path now resolves. Both must fire.
   {:id "the gmail manifest names a different component"
    :edit [:sub (path/join G "kotodama.jsonld")
           "\"path\": \"component.wasm\"" "\"path\": \"gmail-mcp-component\""]
    :must #{:gmail-manifest-component-path :gmail-manifest-component-is-absent}}

   {:id "the wasm component the gmail manifest names actually appears"
    :edit [:write (path/join G "component.wasm") "not-really-wasm"]
    :must #{:gmail-manifest-component-is-absent}}

   {:id "the committed gmail binary becomes a wasm module"
    :edit [:magic gmail-bin]
    :must #{:gmail-binary-is-mach-o-not-wasm}}

   {:id "one mangled gmail NSID is repaired"
    :edit [:sub (path/join G "kotodama.jsonld")
           "emailUserviceUadapter.email_service_adapter_entity"
           "emailServiceAdapter.email_service_adapter_entity"]
    :must #{:gmail-collections-are-mangled}}

   {:id "the declared design system is actually imported"
    :edit [:sub (path/join O "svelte" "src" "App.svelte")
           "import { onMount } from \"svelte\";"
           "import { onMount } from \"svelte\";\n  import \"@etzhayyim/design-system\";"]
    :must #{:design-system-never-imported}}

   {:id "a workspace root appears, so workspace:* could resolve"
    :edit [:write "pnpm-workspace.yaml" "packages:\n  - 'appview/*/svelte'\n"]
    :must #{:workspace-protocol-with-no-workspace-root}}])

;; This machine runs many agents and has repeatedly sat at 100% disk. The binary
;; is 6.7 MB and twelve real copies of it are 80 MB of pure waste, so it is
;; hard-linked instead -- size and magic bytes read identically through a link.
;; The one mutation that rewrites it therefore must NOT write in place, or it
;; would corrupt the checked-in file through the link. It unlinks first, and
;; `original-sha` below is the seatbelt on that reasoning rather than a comment
;; asserting it.
(defn- copy-tree! [src dst]
  (fs/cpSync src dst
             #js {:recursive true
                  :filter (fn [s _]
                            (let [b (path/basename s)]
                              (and (not= b ".git")
                                   (not= b "node_modules")
                                   (not (str/ends-with? s gmail-bin)))))})
  (fs/linkSync (path/join src gmail-bin) (path/join dst gmail-bin)))

(defn- apply-edit! [dir [op rel a b]]
  (let [p (path/join dir rel)]
    (case op
      :sub   (let [s (str (fs/readFileSync p "utf8"))]
               (when-not (str/includes? s a)
                 (throw (js/Error. (str "mutation anchor not found in " rel ": " a))))
               (fs/writeFileSync p (str/replace s a b)))
      :write (do (fs/mkdirSync (path/dirname p) #js {:recursive true})
                 (fs/writeFileSync p a))
      ;; unlink first, then write a same-sized buffer whose first four bytes are
      ;; the wasm magic, so only the magic check moves and the size check holds.
      :magic (let [n (.-size (fs/statSync p))
                   buf (js/Buffer.alloc n)]
               (doseq [[i byte] (map-indexed vector [0x00 0x61 0x73 0x6d])]
                 (aset buf i byte))
               (fs/unlinkSync p)
               (fs/writeFileSync p buf)))))

(defn- sha256 [p]
  (str (cp/execSync (str "shasum -a 256 " (pr-str p) " | cut -d' ' -f1")
                    #js {:encoding "utf8"})))

(defn- failed-ids
  "The check ids the verifier reported as FAIL."
  [out]
  (->> (str/split-lines out)
       (keep #(second (re-find #"^\s*FAIL\s+(\S+)\s*$" %)))
       (map keyword)
       set))

(defn- run-verifier [dir]
  (try
    {:exit 0 :out (str (cp/execSync (str "nbb " (pr-str verifier) " --root " (pr-str dir))
                                    #js {:encoding "utf8" :stdio "pipe"}))}
    (catch :default e
      {:exit (or (.-status e) 1)
       :out (str (some-> (.-stdout e) str) (some-> (.-stderr e) str))})))

;; ---------------------------------------------------------------------------

(def original-sha (sha256 (path/join root gmail-bin)))
(def base (fs/mkdtempSync (path/join (os/tmpdir) "surface-mut-")))
(println (str "scratch\t" base))
(println)

;; Control: the unmutated copy must be green, or nothing below means anything.
(def control (path/join base "control"))
(copy-tree! root control)
(let [{:keys [exit]} (run-verifier control)]
  (when-not (zero? exit)
    (println "CONTROL IS NOT GREEN — an unmutated copy exits" exit)
    (println "Nothing below can demonstrate anything. Stopping.")
    (js/process.exit 1))
  (println "  ok    control (unmutated copy) exits 0"))
(fs/rmSync control #js {:recursive true :force true})
(println)

(def results
  (doall
   (for [[i c] (map-indexed vector cases)]
     (let [dir (path/join base (str "m" i))]
       (copy-tree! root dir)
       (apply-edit! dir (:edit c))
       (let [{:keys [exit out]} (run-verifier dir)
             got (failed-ids out)
             ok (and (= exit 1) (= got (:must c)))]
         (fs/rmSync dir #js {:recursive true :force true})
         (println (str (if ok "  ok  " "  FAIL") "  " (:id c)))
         (when-not ok
           (println (str "        exit:     " exit " (wanted 1)"))
           (println (str "        wanted:   " (pr-str (vec (sort (:must c))))))
           (println (str "        reported: " (pr-str (vec (sort got)))))
           (when (str/includes? out "COULD NOT ANSWER")
             (println "        (verifier answered 2 — the mutation removed an input")
             (println "         rather than changing a fact; not a demonstration)")))
         {:ok ok :id (:id c)})))))

(fs/rmSync base #js {:recursive true :force true})

;; The hard-link hazard, checked rather than asserted.
(when-not (= original-sha (sha256 (path/join root gmail-bin)))
  (println)
  (println "FAIL\tthe checked-in gmail binary changed during this run.")
  (println "\tA mutation wrote through the hard link. Restore it from git.")
  (js/process.exit 1))

(println)
(let [bad (remove :ok results)]
  (if (seq bad)
    (do (println (str "FAIL\t" (count bad) " of " (count results)
                      " mutations were not caught by the intended check."))
        (js/process.exit 1))
    (do (println (str "OK\t" (count results) " mutations, each caught by exactly the check"
                      " that pins it;"))
        (println "\tthe unmutated control stays green and the committed binary is untouched.")
        (js/process.exit 0))))
