from flask import Flask, request, jsonify, send_from_directory, send_file, render_template
import json, os, math, uuid
from datetime import datetime

app = Flask(__name__, static_folder=".", static_url_path="", template_folder="templates")

SUBMISSIONS_FILE  = "parent_reviews_submissions.json"
GARDENS_FILE      = "gardens.json"
PARENTS_FILE      = "Parents_Input.xlsx"
MATCHING_FILE     = "Gale_Shapley_Matching_Result.xlsx"
ADMIN_LOG_FILE    = "admin_run_log.json"

# ─── עזר ───────────────────────────────────────────────────────────────────────
def load_json(path, default):
    return json.load(open(path, encoding='utf-8')) if os.path.exists(path) else default

def save_json(path, data):
    json.dump(data, open(path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

def load_gardens():
    data = load_json(GARDENS_FILE, [])
    return data if isinstance(data, list) else data.get('gardens', [])

def load_parents():
    try:
        import pandas as pd
        if not os.path.exists(PARENTS_FILE):
            return []
        df = pd.read_excel(PARENTS_FILE)
        # המרת NaN לערכים ריקים כדי שה-JSON יעבוד תקין
        df = df.where(pd.notna(df), None)
        return df.to_dict(orient='records')
    except Exception:
        return []

def append_parent(row):
    import pandas as pd
    new = pd.DataFrame([row])
    if os.path.exists(PARENTS_FILE):
        updated = pd.concat([pd.read_excel(PARENTS_FILE), new], ignore_index=True)
    else:
        updated = new
    updated.to_excel(PARENTS_FILE, index=False)

# ─── דפי HTML (Jinja2) ─────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/search-results')
def search_results():
    return render_template('index.html')

@app.route('/map')
def map_page():
    return render_template('map.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/garden-profile')
def garden_profile():
    return render_template('garden-profile.html')

@app.route('/add-review')
def add_review():
    return render_template('add-review.html')

@app.route('/parent-registration')
def parent_registration():
    return render_template('parent-registration.html')

@app.route('/admin')
def admin():
    return render_template('admin.html')

# ─── קבצים סטטיים (JSON, תמונות, CSS) ────────────────────────────────────────
@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory('.', filename)

# ─── ביקורות (קיים) ────────────────────────────────────────────────────────────
def garden_exists(gid):
    return any(str(g.get('id')) == str(gid) for g in load_gardens())

@app.route('/submit_review', methods=['POST'])
def submit_review():
    data = request.get_json()
    if not data:
        return jsonify({'success': False, 'message': 'No data'}), 400
    gid = data.get('garden_id')
    if not gid:
        return jsonify({'success': False, 'message': 'Missing garden_id'}), 400
    if not garden_exists(gid):
        return jsonify({'success': False, 'message': 'Garden not found'}), 404
    subs = load_json(SUBMISSIONS_FILE, [])
    data.update({'submission_id': len(subs)+1,
                 'server_created_at': datetime.now().isoformat(timespec='seconds'),
                 'status': 'pending_review', 'source': 'website_form'})
    subs.append(data)
    save_json(SUBMISSIONS_FILE, subs)
    return jsonify({'success': True, 'submission_id': data['submission_id']})

@app.route('/admin/submissions')
def view_submissions():
    subs = load_json(SUBMISSIONS_FILE, [])
    return jsonify({'total_submissions': len(subs), 'submissions': subs})

# ─── הצעת גנים ─────────────────────────────────────────────────────────────────
@app.route('/suggest_gardens', methods=['POST'])
def suggest_gardens():
    p = request.get_json() or {}
    gardens = load_gardens()
    home_n = (p.get('home_neighborhood') or '').strip()
    needs_tz = p.get('needs_tzaharon') == 'כן'
    imp_tz = int(p.get('importance_tzaharon') or 3)
    pref_edu = (p.get('preferred_education_type') or '').strip()
    pref_rel = (p.get('preferred_religious_orientation') or '').strip()
    pref_gen = (p.get('preferred_gender_composition') or '').strip()
    imp_gen = int(p.get('importance_gender_composition') or 1)
    pref_lang = (p.get('preferred_activity_language') or '').strip()
    needs_prot = p.get('needs_protected_space') == 'כן'
    needs_fri = p.get('needs_friday') == 'כן'
    max_price = p.get('max_price')
    child_age = p.get('child_age_months')
    try: child_age = int(child_age)
    except (TypeError, ValueError): child_age = None

    # אם ההורה ביקש סוג חינוך ספציפי — מחפשים בכל העיר, אחרת — שכונה קודם
    city_wide_search = bool(pref_edu)

    scored = []
    for g in gardens:
        score = 0
        g_n = (g.get('neighborhood') or '').strip()
        g_sector = (g.get('sector') or '').strip()
        g_edu = (g.get('education_type') or '').strip()
        has_tz = g.get('has_tzaharon') == 'כן'
        has_fri = g.get('friday') == 'כן'
        has_prot = bool(g.get('has_protected_space'))
        g_gen = (g.get('gender_composition') or '').strip()
        price_avg = g.get('price_avg')
        g_min_age = g.get('min_age_months')
        g_max_age = g.get('max_age_months')

        # סינון גיל קשיח
        if child_age is not None and g_min_age is not None and g_max_age is not None:
            try:
                if not (int(g_min_age) <= child_age < int(g_max_age)):
                    continue
            except (TypeError, ValueError):
                pass

        # סינון שכונה: אם לא מחפשים בכל העיר — רק שכונת ההורה
        if not city_wide_search and home_n and g_n and g_n != home_n:
            continue

        # סינון חינוך קשיח כשיש העדפה
        if pref_edu and g_edu and g_edu != pref_edu:
            continue

        # סינון צהרון קשיח
        if needs_tz and imp_tz >= 4 and not has_tz:
            continue

        # סינון מגדר קשיח
        if pref_gen and imp_gen >= 4 and g_gen and g_gen not in ('', pref_gen, 'מעורב'):
            continue

        # ניקוד
        if home_n and g_n == home_n: score += 40
        if pref_edu and g_edu == pref_edu: score += 30
        if pref_rel:
            g_rel = (g.get('religious_orientation') or g_sector or '').strip()
            if pref_rel in g_rel or g_rel in pref_rel: score += 25
        if pref_gen and g_gen:
            score += 20 if pref_gen == g_gen else (5 if g_gen == 'מעורב' else 0)
        if pref_lang and (g.get('activity_language') or '').strip() == pref_lang: score += 15
        if needs_tz and has_tz: score += 20
        if needs_prot and has_prot: score += 10
        if needs_fri and has_fri: score += 10
        if max_price and price_avg:
            try:
                if float(price_avg) <= float(max_price): score += 15
            except (ValueError, TypeError): pass

        age_label = ''
        if g_min_age is not None and g_max_age is not None:
            age_label = f'{g_min_age}–{g_max_age} חודשים'

        scored.append({
            'id': str(g.get('id','')),
            'name': g.get('display_name') or g.get('name',''),
            'address': g.get('address',''), 'neighborhood': g_n,
            'sector': g_sector, 'education_type': g_edu,
            'has_tzaharon': 'כן' if has_tz else 'לא',
            'friday': 'כן' if has_fri else 'לא',
            'price_avg': price_avg or '',
            'age_label': age_label,
            'score': score,
            'distance_label': 'באותה שכונה' if home_n and g_n == home_n else ''
        })

    scored.sort(key=lambda x: -x['score'])
    return jsonify({'gardens': scored[:30], 'total_filtered': len(scored)})

# ─── הרשמת הורה ────────────────────────────────────────────────────────────────
@app.route('/register_parent', methods=['POST'])
def register_parent():
    data = request.get_json() or {}
    for f in ['parent_name','child_name','home_address']:
        if not data.get(f):
            return jsonify({'success': False, 'message': f'Missing: {f}'}), 400
    pid = 'P' + datetime.now().strftime('%Y%m%d%H%M%S') + str(uuid.uuid4())[:4].upper()
    row = {
        'parent_id': pid, 'parent_name': data.get('parent_name',''),
        'parent_phone': data.get('parent_phone',''), 'child_name': data.get('child_name',''),
        'child_age_months': data.get('child_age_months',''), 'child_gender': data.get('child_gender',''),
        'child_birth_month': data.get('child_birth_month',''),
        'home_address': data.get('home_address',''), 'home_neighborhood': data.get('home_neighborhood',''),
        'home_lat': '', 'home_lon': '',
        'max_distance_km': data.get('max_distance_km', 2.5), 'max_price': data.get('max_price',''),
        'preferred_activity_language': data.get('preferred_activity_language',''),
        'preferred_education_type': data.get('preferred_education_type',''),
        'preferred_religious_orientation': data.get('preferred_religious_orientation',''),
        'preferred_religious_substream': '', 'preferred_gender_composition': data.get('preferred_gender_composition',''),
        'preferred_sector': '', 'needs_friday': data.get('needs_friday','לא'),
        'needs_protected_space': data.get('needs_protected_space','לא'),
        'needs_tzaharon': data.get('needs_tzaharon','לא'),
        'declared_sibling_in_garden': '', 'sibling_garden_id': '', 'sibling_verification_status': '',
        'preferred_garden_1': data.get('preferred_garden_1',''), 'preferred_garden_2': data.get('preferred_garden_2',''),
        'preferred_garden_3': data.get('preferred_garden_3',''),
        'preferred_garden_1_name': data.get('preferred_garden_1_name',''),
        'preferred_garden_2_name': data.get('preferred_garden_2_name',''),
        'preferred_garden_3_name': data.get('preferred_garden_3_name',''),
        'importance_distance': data.get('importance_distance',3), 'importance_price': data.get('importance_price',3),
        'importance_activity_language': data.get('importance_activity_language',3),
        'importance_education_type': data.get('importance_education_type',3),
        'importance_religious_orientation': data.get('importance_religious_orientation',3),
        'importance_religious_substream': 3,
        'importance_gender_composition': data.get('importance_gender_composition',1),
        'importance_sector': 3, 'importance_friday': 3, 'importance_protected_space': 3,
        'importance_tzaharon': data.get('importance_tzaharon',3),
        'friend_request_1': '', 'friend_request_2': '', 'importance_same_friend': 3,
        'willing_to_trade_distance_for_friend': 'לא', 'allow_manual_far_preference': 'לא',
        'max_manual_exception_distance_km': '',
        'registered_at': datetime.now().isoformat(timespec='seconds'),
    }
    try:
        append_parent(row)
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500
    return jsonify({'success': True, 'parent_id': pid})

# ─── ממשק מנהל ─────────────────────────────────────────────────────────────────
@app.route('/admin/stats')
def admin_stats():
    log = load_json(ADMIN_LOG_FILE, {})
    return jsonify({'parents_count': len(load_parents()), 'gardens_count': len(load_gardens()),
                    'last_run': log.get('last_run'), 'last_status': log.get('last_status','—')})

@app.route('/admin/parents')
def admin_parents():
    parents = load_parents()
    return jsonify({'parents': parents, 'total': len(parents)})

@app.route('/run_matching', methods=['POST'])
def run_matching():
    if not os.path.exists(PARENTS_FILE):
        return jsonify({'success': False, 'message': 'אין הורים רשומים'}), 400
    try:
        import pandas as pd
        from gale_shapley_matching import (
            ensure_gardens_input_file, validate_parents_input, clean_value, to_int,
            DEFAULT_CAPACITY, build_preferences, gale_shapley_with_capacities,
            mandatory_nearby_assignment, apply_manual_exception_after_all_close_solution,
            build_stability_check, build_distance_audit, build_preference_satisfaction, create_outputs,
        )
        gardens_df = ensure_gardens_input_file()
        parents_df = pd.read_excel(PARENTS_FILE)
        validate_parents_input(parents_df)
        parents_df['parent_id'] = parents_df['parent_id'].apply(clean_value)
        gardens_df['garden_id'] = gardens_df['garden_id'].apply(clean_value)
        capacities = {clean_value(r.get('garden_id','')): to_int(r.get('capacity', DEFAULT_CAPACITY), default=DEFAULT_CAPACITY)
                      for _, r in gardens_df.iterrows() if clean_value(r.get('garden_id',''))}
        (parent_prefs, garden_ranks, score_lookup, garden_priority_lookup,
         parent_prefs_df, garden_ranks_df, candidate_df, social_df) = build_preferences(parents_df, gardens_df)
        parent_match, garden_matches, proposal_df = gale_shapley_with_capacities(parent_prefs, garden_ranks, capacities)
        assignment_source = {pid: 'gale_shapley' for pid, gid in parent_match.items() if gid}
        # בדיקת יציבות על תוצאת GS הטהורה — לפני כל שלב ידני
        stability_df, blocking_df = build_stability_check(
            parent_prefs, garden_ranks, parent_match, garden_matches, capacities, parents_df, gardens_df)
        mandatory_df, planning_df = mandatory_nearby_assignment(
            parents_df, gardens_df, parent_match, garden_matches, capacities, score_lookup, garden_priority_lookup)
        for _, r in planning_df.iterrows():
            assignment_source[clean_value(r.get('parent_id',''))] = 'requires_planning_review'
        manual_df = apply_manual_exception_after_all_close_solution(
            parents_df, gardens_df, parent_match, garden_matches, capacities,
            score_lookup, garden_priority_lookup, pd.DataFrame())
        distance_df = build_distance_audit(parents_df, gardens_df, parent_match, assignment_source)
        pref_sat_df = build_preference_satisfaction(parents_df, gardens_df, parent_match, assignment_source)
        create_outputs(parents_df, gardens_df, parent_match, garden_matches,
                       score_lookup, garden_priority_lookup, parent_prefs_df, garden_ranks_df,
                       candidate_df, proposal_df, stability_df, blocking_df, capacities,
                       mandatory_df, planning_df, manual_df, distance_df, assignment_source, social_df)

        gd = {clean_value(r.get('garden_id','')): r.to_dict() for _, r in gardens_df.iterrows()}
        total = len(parents_df)
        matched = sum(1 for v in parent_match.values() if v)
        blocking = len(blocking_df) if blocking_df is not None else 0

        final = []
        for _, r in parents_df.iterrows():
            pid = clean_value(r.get('parent_id',''))
            gid = parent_match.get(pid)
            grow = gd.get(gid, {})
            final.append({'parent_name': clean_value(r.get('parent_name','')),
                          'child_name': clean_value(r.get('child_name','')),
                          'garden_name': clean_value(grow.get('garden_name', gid or '—')),
                          'garden_address': clean_value(grow.get('address','')),
                          'parent_fit_score': score_lookup.get(pid,{}).get(gid,''),
                          'status': 'matched' if gid else 'unmatched'})

        cap_rows = sorted([{'garden_name': clean_value(gd.get(g,{}).get('garden_name',g)),
                             'capacity': c, 'assigned': len(garden_matches.get(g,[])),
                             'remaining': c - len(garden_matches.get(g,[]))}
                            for g, c in capacities.items()], key=lambda x: -x['assigned'])

        pref_rows = []
        for _, r in parents_df.iterrows():
            pid = clean_value(r.get('parent_id',''))
            gid = parent_match.get(pid)
            def _cg(v):
                try: return str(int(float(v))) if v not in (None,'') else ''
                except: return clean_value(v)
            p1,p2,p3 = [_cg(r.get(f'preferred_garden_{i}','')) for i in [1,2,3]]
            rank = 1 if gid==p1 else (2 if gid==p2 else (3 if gid==p3 else None))
            grow = gd.get(gid,{})
            pref_rows.append({'parent_name': clean_value(r.get('parent_name','')),
                               'garden_name': clean_value(grow.get('garden_name', gid or '—')),
                               'preference_rank': rank})

        proposal_count = len(proposal_df) if proposal_df is not None else 0

        garden_demand = {}
        for glist in parent_prefs.values():
            for gid in glist:
                garden_demand[gid] = garden_demand.get(gid, 0) + 1
        demand_rows = sorted([{'garden_name': clean_value(gd.get(g,{}).get('garden_name',g)), 'demand': cnt}
                               for g, cnt in garden_demand.items()], key=lambda x: -x['demand'])
        no_assign_rows = [{'garden_name': clean_value(gd.get(g,{}).get('garden_name',g)), 'capacity': c}
                          for g, c in capacities.items() if not garden_matches.get(g)]

        save_json(ADMIN_LOG_FILE, {'last_run': datetime.now().isoformat(timespec='seconds'),
                                    'last_status': 'יציב ✓' if blocking==0 else f'{blocking} זוגות מערערים'})
        return jsonify({'success': True,
            'summary': {'parents_total': total, 'parents_matched_total': matched,
                        'matched_by_gale_shapley': sum(1 for s in assignment_source.values() if s=='gale_shapley'),
                        'matched_by_mandatory_nearby_assignment': sum(1 for s in assignment_source.values() if s=='mandatory_nearby_assignment'),
                        'requires_planning_review': len(planning_df),
                        'gardens_total': len(gardens_df), 'gardens_used': sum(1 for v in garden_matches.values() if v),
                        'total_capacity': sum(capacities.values()),
                        'blocking_pairs_found': blocking, 'stable_matching': 'yes' if blocking==0 else 'no',
                        'average_candidates_per_parent': round(sum(len(v) for v in parent_prefs.values())/max(total,1),2),
                        'proposal_count': proposal_count},
            'final_matching': final, 'capacity_utilization': cap_rows,
            'preference_satisfaction': pref_rows,
            'stability': {'stable': blocking==0, 'blocking_pairs_found': blocking},
            'garden_demand': demand_rows,
            'gardens_no_assignment': no_assign_rows})
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'message': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/download_matching_result')
def download_result():
    if not os.path.exists(MATCHING_FILE):
        return jsonify({'error': 'קובץ לא קיים'}), 404
    return send_file(MATCHING_FILE, as_attachment=True, download_name='Gale_Shapley_Matching_Result.xlsx')

if __name__ == '__main__':
    app.run(debug=True)