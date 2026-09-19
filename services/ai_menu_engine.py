"""團膳 AI 菜單規則引擎：排菜、評分、換菜與結果重算。"""
from __future__ import annotations
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
import random
import re
import secrets
from sqlalchemy.orm import joinedload
from models import KitchenRecipe, KitchenRecipeIngredient
from services.nutrition import recipe_nutrition
from services.food_servings import GROUP_ORDER, finalize_day_servings, merge_recipe_servings, recipe_servings, serving_deviation, targets_for_kcal
from services.recipe_tags import effective_rule_tags

CATEGORY_ORDER=("主食","主菜","副菜","青菜","湯品","點心")
DEFAULT_STRUCTURE={"主食":1,"主菜":1,"副菜":2,"青菜":1,"湯品":1,"點心":0}
STRUCTURE_LIMITS={"主食":2,"主菜":2,"副菜":4,"青菜":2,"湯品":2,"點心":2}
HARD_RULE_PENALTY=1_000_000
WEEKDAY_LABEL="一二三四五六日"
CAP_TAGS=("fish","fried","sweet_soup")
CAP_LABEL={"fish":"魚類","fried":"炸物","sweet_soup":"甜湯"}


class RuleFeasibilityError(ValueError):
    """Raised when a requested menu cannot satisfy a hard weekly rule."""


def _as_int(value, default=0):
    try:return max(0,int(value))
    except (TypeError,ValueError):return default

def parse_weekdays(value)->tuple[int,...]:
    if value in (None,"",[]):return ()
    if isinstance(value,str):
        value=[part.strip() for part in value.replace("，",",").split(",") if part.strip()]
    out=[]
    for item in value:
        try:n=int(item)
        except (TypeError,ValueError):continue
        if 0<=n<=6:out.append(n)
    return tuple(sorted(set(out)))

def weekday_text(days)->str:
    if not days:return "不限星期"
    return "、".join(f"週{WEEKDAY_LABEL[d]}" for d in days)

def _cap_max(rules,tag):
    return getattr(rules,f"{tag}_per_week_max")

def _cap_days(rules,tag):
    return getattr(rules,f"{tag}_weekdays") or ()


@dataclass
class Rules:
    recipe_repeat_days:int=5
    main_repeat_days:int=5
    fish_per_week_min:int=0
    fish_per_week_max:int=1
    fried_per_week_max:int=1
    sweet_soup_per_week_max:int=1
    fish_weekdays:tuple[int,...]=()
    fried_weekdays:tuple[int,...]=()
    sweet_soup_weekdays:tuple[int,...]=()
    prefer_kcal_in_range:bool=True
    exclude_incomplete_nutrition:bool=True
    @classmethod
    def from_dict(cls,data):
        data=data or {}; r=cls()
        legacy=data.get("repeat_days")
        if legacy is not None:r.recipe_repeat_days=r.main_repeat_days=_as_int(legacy)
        for k in ("recipe_repeat_days","main_repeat_days","fish_per_week_min","fish_per_week_max","fried_per_week_max","sweet_soup_per_week_max"):
            if data.get(k) is not None:setattr(r,k,_as_int(data[k],getattr(r,k)))
        for k in ("fish_weekdays","fried_weekdays","sweet_soup_weekdays"):
            if k in data:setattr(r,k,parse_weekdays(data.get(k)))
        for k in ("prefer_kcal_in_range","exclude_incomplete_nutrition"):
            if k in data:setattr(r,k,bool(data[k]))
        return r
    def to_dict(self):
        data=dict(self.__dict__)
        for k in ("fish_weekdays","fried_weekdays","sweet_soup_weekdays"):
            data[k]=list(getattr(self,k) or ())
        return data

@dataclass(frozen=True)
class Candidate:
    id:int; name:str; category:str; kcal:int|None; tags:frozenset[str]=frozenset(); servings:tuple=(); vegetarian_compatible:bool=False
    def serving_dict(self):return dict(self.servings)

@dataclass
class DayResult:
    service_date:date; kcal:int|None; warnings:list[str]=field(default_factory=list); ok_messages:list[str]=field(default_factory=list); servings:dict[str,float]=field(default_factory=dict)

@dataclass
class PlanResult:
    dates:list[date]; structure:dict; rules:Rules; assignment:dict; candidates:dict[int,Candidate]; penalty:float; breakdown:dict[str,float]; days:dict[date,DayResult]; warnings:list[str]; kcal_min:int|None=None; kcal_max:int|None=None; meal_variant:str="regular"
    def recipe_at(self,day_index,category,slot_index):
        rid=self.assignment.get((day_index,category,slot_index));return self.candidates.get(rid) if rid else None

def service_dates(start,end,include_weekends=False):
    out=[];d=start
    while d<=end:
        if include_weekends or d.weekday()<5:out.append(d)
        d+=timedelta(days=1)
    return out

def build_structure(values):
    values=values or {};out={}
    for c in CATEGORY_ORDER:
        try:n=int(values.get(c,DEFAULT_STRUCTURE[c]))
        except (TypeError,ValueError):n=DEFAULT_STRUCTURE[c]
        out[c]=max(0,min(STRUCTURE_LIMITS[c],n))
    return out

def describe_rules(r):
    lines=[f"同一道菜 {r.recipe_repeat_days} 天內不重複",f"主菜 {r.main_repeat_days} 天內不重複"]
    for tag in CAP_TAGS:
        lines.append(f"{CAP_LABEL[tag]}每週最多 {_cap_max(r,tag)} 次，{weekday_text(_cap_days(r,tag))}出現（硬限制）")
    if r.fish_per_week_min:lines.append(f"每週至少 {r.fish_per_week_min} 次魚類（硬限制）")
    if r.prefer_kcal_in_range:lines.append("每日總熱量與六大類份數盡量接近供餐對象基準")
    if r.exclude_incomplete_nutrition:lines.append("排除營養資料不完整的菜色")
    return lines

def _vegetarian_compatible(recipe,tags):
    ingredient_names=[row.ingredient.name or "" for row in recipe.ingredients or () if row.ingredient]
    if not ingredient_names:return False
    names=[recipe.name or ""]+ingredient_names
    text=" ".join(names).replace("雞蛋","蛋")
    for phrase in ("素魚排","素雞排","素肉排","素火腿","素黑輪","素鍋貼","素燒賣","素水餃","素甜不辣","素肉","素魚","素排","素雞"):text=text.replace(phrase," ")
    animal_terms=("豬","牛","羊","雞","鴨","鵝","魚","蝦","蟹","蛤","貝","蚵","魷","花枝","小卷","海鮮","肉","排骨","大骨","骨湯","火腿","培根","貢丸","肉羹","魚羹","黑輪","鍋貼","燒賣","水餃","甜不辣","熱狗","香腸","米血","血糕")
    ambiguous_terms=("丸","排","腿","翅","捲","卷","堡","羹","獅子頭","滷味","關東煮","沙茶","柴魚","蠔油")
    if any(term in text for term in animal_terms):return False
    if any(term in text for term in ambiguous_terms):return False
    # 這四道正式配方雖未列肉類，但實際作業不是素食，不進共用池。
    if (recipe.name or "").strip() in {"羅宋湯","玉米濃湯","酸辣湯","招牌炒飯"}:return False
    return True

def _vegetarian_pair_score(candidate,reference):
    if not reference:return 0
    if candidate.id==reference.id:return -10000
    normalize=lambda name:(name or "").replace("(素)","").replace("（素）","").replace("素食","").strip()
    return -9000 if normalize(candidate.name)==normalize(reference.name) else 0

def load_candidates():
    rows=(KitchenRecipe.query.options(joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient),joinedload(KitchenRecipe.tags)).filter(KitchenRecipe.active.is_(True)).order_by(KitchenRecipe.name).all())
    out=[]
    for r in rows:
        cat=(r.category or "").strip()
        if cat not in CATEGORY_ORDER:continue
        n=recipe_nutrition(r);tags=frozenset(effective_rule_tags(r));sv=recipe_servings(r).rounded()
        out.append(Candidate(r.id,r.name,cat,n.kcal_int,tags,tuple((k,sv.get(k,0.0)) for k in GROUP_ORDER),_vegetarian_compatible(r,tags)))
    return out

def _candidate_map(candidates):
    return candidates if isinstance(candidates,dict) else {c.id:c for c in candidates}

def candidate_allowed(candidate,meal_variant):
    if not candidate:return False
    if meal_variant=="vegetarian":return candidate.vegetarian_compatible
    # A labelled recipe stays vegetarian-facing; regular menus use its
    # unlabelled twin with the same BOM.
    if meal_variant=="regular":return not bool(re.search(r"[（(]素[）)]\s*$",candidate.name or ""))
    return True

def variant_violations(assignment,candidates,meal_variant):
    by_id=_candidate_map(candidates);bad=sorted({by_id[rid].name for rid in assignment.values() if rid in by_id and not candidate_allowed(by_id[rid],meal_variant)})
    return [f"{('素食' if meal_variant=='vegetarian' else '葷食')}菜單含不相容菜色：{'、'.join(bad)}"] if bad else []

def _week_tag_counts(dates,assignment,candidates):
    by_id=_candidate_map(candidates);counts={d.isocalendar()[:2]:Counter() for d in dates}
    for key,rid in assignment.items():
        if not rid:continue
        i=key[0]
        if i<0 or i>=len(dates):continue
        cand=by_id.get(rid)
        if not cand:continue
        week=dates[i].isocalendar()[:2]
        for tag in cand.tags:counts.setdefault(week,Counter())[tag]+=1
    return counts

def _day_tags(dates,assignment,candidates,day_index):
    by_id=_candidate_map(candidates);tags=set()
    for key,rid in assignment.items():
        if not rid or key[0]!=day_index:continue
        cand=by_id.get(rid)
        if cand:tags.update(cand.tags)
    return tags

def _required_tag_dates(dates,weekdays,maxn):
    if not weekdays or maxn<=0:return []
    hits=[d for d in dates if d.weekday() in weekdays]
    return [] if len(hits)>maxn else hits

def _hard_cap_violations(dates,rules,assignment,candidates):
    violations=[]
    by_id=_candidate_map(candidates)
    for week,cnt in _week_tag_counts(dates,assignment,candidates).items():
        for tag in CAP_TAGS:
            maxn=_cap_max(rules,tag)
            if cnt[tag]>maxn:
                violations.append(f"{week[0]}年第{week[1]}週{CAP_LABEL[tag]} {cnt[tag]} 次，超過上限 {maxn} 次")
    for key,rid in assignment.items():
        if not rid:continue
        i=key[0]
        if i<0 or i>=len(dates):continue
        cand=by_id.get(rid)
        if not cand:continue
        weekday=dates[i].weekday()
        for tag in CAP_TAGS:
            allowed=_cap_days(rules,tag)
            if allowed and tag in cand.tags and weekday not in allowed:
                violations.append(f"{dates[i]:%Y-%m-%d} 出現{CAP_LABEL[tag]}，但規則只允許 {weekday_text(allowed)} 出現")
    return violations

def hard_rule_violations(dates,rules,assignment,candidates):
    violations=list(_hard_cap_violations(dates,rules,assignment,candidates))
    for week,cnt in _week_tag_counts(dates,assignment,candidates).items():
        if cnt["fish"]<rules.fish_per_week_min:
            violations.append(f"{week[0]}年第{week[1]}週魚類 {cnt['fish']} 次，未達至少 {rules.fish_per_week_min} 次")
    index={d:i for i,d in enumerate(dates)}
    for tag in CAP_TAGS:
        for day in _required_tag_dates(dates,_cap_days(rules,tag),_cap_max(rules,tag)):
            if tag not in _day_tags(dates,assignment,candidates,index[day]):
                violations.append(f"{day:%Y-%m-%d}（週{WEEKDAY_LABEL[day.weekday()]}）應出現{CAP_LABEL[tag]}，但當天沒有")
    return violations

def evaluate(dates,structure,rules,assignment,candidates,kcal_min=None,kcal_max=None):
    by_id={c.id:c for c in candidates};breakdown=defaultdict(float);day_results={};warnings=[];uses=defaultdict(list);main_uses=defaultdict(list);week_tags={d.isocalendar()[:2]:Counter() for d in dates}
    for i,d in enumerate(dates):
        total=0;complete=True;wrn=[];ok=[];seen=set()
        for c in CATEGORY_ORDER:
            for s in range(1,structure.get(c,0)+1):
                rid=assignment.get((i,c,s));cand=by_id.get(rid) if rid else None
                if not cand:complete=False;breakdown["empty"]+=20000;continue
                if rid in seen:breakdown["duplicate"]+=20000
                seen.add(rid);uses[rid].append(i)
                if c=="主菜":main_uses[rid].append(i)
                if cand.kcal is None:complete=False;wrn.append(f"⚠ {cand.name} 營養資料不完整")
                else:total+=cand.kcal
                for tag in cand.tags:week_tags[d.isocalendar()[:2]][tag]+=1
        kcal=total if complete else None
        base_servings=merge_recipe_servings(by_id[rid].serving_dict() for key,rid in assignment.items() if key[0]==i and rid in by_id)
        servings=finalize_day_servings(base_servings,kcal)
        if kcal is not None and kcal_min is not None and kcal_max is not None:
            if kcal<kcal_min:
                delta=kcal_min-kcal;breakdown["kcal"]+=min(10000,delta*25);wrn.append(f"⚠ 熱量 {kcal} kcal 低於基準 {kcal_min}～{kcal_max} kcal")
            elif kcal>kcal_max:
                delta=kcal-kcal_max;breakdown["kcal"]+=min(10000,delta*25);wrn.append(f"⚠ 熱量 {kcal} kcal 高於基準 {kcal_min}～{kcal_max} kcal")
            else:ok.append(f"✅ 熱量 {kcal} kcal 符合基準")
        elif kcal is None:wrn.append("⚠ 無法計算當日總熱量")
        else:ok.append(f"總熱量 {kcal} kcal")
        targets=targets_for_kcal(kcal_min,kcal_max)
        if targets:
            deviation,serving_warnings,serving_ok=serving_deviation(servings,targets);breakdown["servings"]+=deviation*1200;wrn.extend(serving_warnings);ok.extend(serving_ok)
        day_results[d]=DayResult(d,kcal,wrn,ok,servings)
    for rid,days_used in uses.items():
        for a,b in zip(days_used,days_used[1:]):
            if rules.recipe_repeat_days and b-a<rules.recipe_repeat_days:breakdown["recipe_repeat"]+=6000
        breakdown["variety"]+=max(0,len(days_used)-1)*60
    for rid,days_used in main_uses.items():
        for a,b in zip(days_used,days_used[1:]):
            if rules.main_repeat_days and b-a<rules.main_repeat_days:breakdown["main_repeat"]+=4000
    for week,cnt in week_tags.items():
        if cnt["fish"]<rules.fish_per_week_min:
            deficit=rules.fish_per_week_min-cnt["fish"];breakdown["hard_fish"]+=deficit*HARD_RULE_PENALTY;warnings.append(f"⛔ {week[0]}年第{week[1]}週：魚類 {cnt['fish']} 次，未達硬限制 {rules.fish_per_week_min} 次")
        for tag in CAP_TAGS:
            maxn=_cap_max(rules,tag)
            if cnt[tag]>maxn:
                excess=cnt[tag]-maxn;breakdown[f"hard_{tag}"]+=excess*HARD_RULE_PENALTY;warnings.append(f"⛔ {week[0]}年第{week[1]}週：{CAP_LABEL[tag]} {cnt[tag]} 次，超過硬限制 {maxn} 次")
    for msg in _hard_cap_violations(dates,rules,assignment,by_id):
        if "不在允許" in msg or "只允許" in msg:
            breakdown["hard_weekday"]+=HARD_RULE_PENALTY;warnings.append(f"⛔ {msg}")
    for msg in hard_rule_violations(dates,rules,assignment,by_id):
        if "應出現" in msg:
            breakdown["hard_weekday"]+=HARD_RULE_PENALTY
            if f"⛔ {msg}" not in warnings:warnings.append(f"⛔ {msg}")
    return sum(breakdown.values()),dict(breakdown),day_results,warnings

def _repair_weekly_fish(dates,structure,rules,assignment,pools,by_id,locked_keys,candidates,kcal_min,kcal_max):
    if rules.fish_per_week_min<=0:return
    weeks=sorted({d.isocalendar()[:2] for d in dates})
    for week in weeks:
        while _week_tag_counts(dates,assignment,by_id)[week]["fish"]<rules.fish_per_week_min:
            best=None
            for key,current_id in list(assignment.items()):
                i,cat,_slot=key
                if key in locked_keys or dates[i].isocalendar()[:2]!=week:continue
                current=by_id.get(current_id)
                if current and "fish" in current.tags:continue
                same={rid for k,rid in assignment.items() if k[0]==i and k!=key and rid}
                for cand in pools.get(cat,[]):
                    if "fish" not in cand.tags or cand.id in same:continue
                    trial=dict(assignment);trial[key]=cand.id
                    if _hard_cap_violations(dates,rules,trial,by_id):continue
                    penalty,_,_,_=evaluate(dates,structure,rules,trial,candidates,kcal_min,kcal_max)
                    score=(penalty,cand.name,key)
                    if best is None or score<best[0]:best=(score,key,cand.id)
            if best is None:
                y,w=week
                raise RuleFeasibilityError(f"{y}年第{w}週無法滿足每週至少 {rules.fish_per_week_min} 次魚類；請增加可用魚類菜色、調整每日結構或降低魚類下限。")
            _score,key,candidate_id=best;assignment[key]=candidate_id

def _weekday_preference(day,candidate,rules):
    score=0;weekday=day.weekday()
    for tag in CAP_TAGS:
        if tag not in candidate.tags:continue
        allowed=_cap_days(rules,tag)
        if allowed and weekday in allowed:score-=400
    return score

def _repair_required_weekdays(dates,structure,rules,assignment,pools,by_id,locked_keys,candidates,kcal_min,kcal_max):
    index={d:i for i,d in enumerate(dates)}
    for tag in CAP_TAGS:
        for day in _required_tag_dates(dates,_cap_days(rules,tag),_cap_max(rules,tag)):
            i=index[day]
            if tag in _day_tags(dates,assignment,by_id,i):continue
            best=None
            for key,current_id in list(assignment.items()):
                day_index,cat,_slot=key
                if key in locked_keys or day_index!=i:continue
                current=by_id.get(current_id)
                if current and tag in current.tags:continue
                same={rid for k,rid in assignment.items() if k[0]==i and k!=key and rid}
                for cand in pools.get(cat,[]):
                    if tag not in cand.tags or cand.id in same:continue
                    trial=dict(assignment);trial[key]=cand.id
                    if _hard_cap_violations(dates,rules,trial,by_id):continue
                    penalty,_,_,_=evaluate(dates,structure,rules,trial,candidates,kcal_min,kcal_max)
                    score=(penalty,cand.name,key)
                    if best is None or score<best[0]:best=(score,key,cand.id)
            if best is None:
                raise RuleFeasibilityError(f"{day:%Y-%m-%d}（週{WEEKDAY_LABEL[day.weekday()]}）無法安排{CAP_LABEL[tag]}；請補充菜色、調整每日結構，或改選其他星期。")
            _score,key,candidate_id=best;assignment[key]=candidate_id

def generate(start,end,structure,rules,kcal_min=None,kcal_max=None,include_weekends=False,seed=None,locked=None,meal_variant="regular",reference_assignment=None):
    if seed is None:
        seed=secrets.randbits(64)
    rng=random.Random(seed);dates=service_dates(start,end,include_weekends);candidates=load_candidates();by_id={c.id:c for c in candidates};pools=defaultdict(list)
    for c in candidates:
        if rules.exclude_incomplete_nutrition and c.kcal is None:continue
        if not candidate_allowed(c,meal_variant):continue
        pools[c.category].append(c)
    for p in pools.values():rng.shuffle(p)
    assignment=dict(locked or {});locked_keys=set(assignment.keys());last={};usage=Counter()
    incompatible=variant_violations(assignment,by_id,meal_variant)
    if incompatible:raise RuleFeasibilityError("；".join(incompatible))
    cap_errors=_hard_cap_violations(dates,rules,assignment,by_id)
    if cap_errors:raise RuleFeasibilityError("；".join(cap_errors))
    for i,d in enumerate(dates):
        day_seen={assignment[k] for k in assignment if k[0]==i and assignment[k]}
        for c in CATEGORY_ORDER:
            for s in range(1,structure.get(c,0)+1):
                key=(i,c,s)
                if key in assignment and assignment[key]:continue
                best=None;bestscore=None
                for cand in pools.get(c,[]):
                    if cand.id in day_seen:continue
                    trial=dict(assignment);trial[key]=cand.id
                    if _hard_cap_violations(dates,rules,trial,by_id):continue
                    gap=i-last.get(cand.id,-999);reference=by_id.get((reference_assignment or {}).get(key));score=usage[cand.id]*60+rng.random()+_vegetarian_pair_score(cand,reference)+_weekday_preference(d,cand,rules)
                    if gap<rules.recipe_repeat_days:score+=6000
                    if c=="主菜" and gap<rules.main_repeat_days:score+=4000
                    if bestscore is None or score<bestscore:best,bestscore=cand,score
                if best is None:
                    raise RuleFeasibilityError(f"{d:%Y-%m-%d} 的「{c}」沒有符合每週硬限制的可用菜色；請補充菜色或放寬魚類／炸物／甜湯上限。")
                assignment[key]=best.id;day_seen.add(best.id);last[best.id]=i;usage[best.id]+=1
    _repair_weekly_fish(dates,structure,rules,assignment,pools,by_id,locked_keys,candidates,kcal_min,kcal_max)
    _repair_required_weekdays(dates,structure,rules,assignment,pools,by_id,locked_keys,candidates,kcal_min,kcal_max)
    violations=variant_violations(assignment,by_id,meal_variant)+hard_rule_violations(dates,rules,assignment,by_id)
    if violations:raise RuleFeasibilityError("產生完成後驗證失敗："+"；".join(violations))
    pair_penalty=lambda values:sum(_vegetarian_pair_score(by_id.get(rid),by_id.get((reference_assignment or {}).get(key))) for key,rid in values.items() if by_id.get(rid))
    current_penalty,_,_,_=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max);current_penalty+=pair_penalty(assignment);movable=[key for key in assignment if key not in locked_keys]
    for _pass in range(2):
        improved=False;rng.shuffle(movable)
        for key in movable:
            i,cat,_slot=key;current=assignment.get(key);same={rid for k,rid in assignment.items() if k[0]==i and k!=key and rid};trial_pool=[cand for cand in pools.get(cat,[]) if cand.id not in same];rng.shuffle(trial_pool);best_id=current;best_penalty=current_penalty
            for cand in trial_pool[:16]:
                if cand.id==current:continue
                trial=dict(assignment);trial[key]=cand.id
                if hard_rule_violations(dates,rules,trial,by_id):continue
                p,_,_,_=evaluate(dates,structure,rules,trial,candidates,kcal_min,kcal_max);p+=pair_penalty(trial)
                if p+0.01<best_penalty:best_id,best_penalty=cand.id,p
            if best_id!=current:assignment[key]=best_id;current_penalty=best_penalty;improved=True
        if not improved:break
    violations=variant_violations(assignment,by_id,meal_variant)+hard_rule_violations(dates,rules,assignment,by_id)
    if violations:raise RuleFeasibilityError("；".join(violations))
    penalty,breakdown,days_report,warnings=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max)
    return PlanResult(dates,structure,rules,assignment,by_id,penalty,breakdown,days_report,warnings,kcal_min,kcal_max,meal_variant)

def rebuild(dates,structure,rules,assignment,kcal_min=None,kcal_max=None,meal_variant="regular"):
    candidates=load_candidates();penalty,breakdown,days_report,warnings=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max);return PlanResult(dates,structure,rules,assignment,{c.id:c for c in candidates},penalty,breakdown,days_report,warnings,kcal_min,kcal_max,meal_variant)

def replacement_options(result,target_key,limit=None):
    i,cat,slot=target_key;same={rid for key,rid in result.assignment.items() if key[0]==i and key!=target_key and rid};current=result.assignment.get(target_key);scored=[]
    for cand in result.candidates.values():
        if cand.category!=cat or cand.id in same or (result.rules.exclude_incomplete_nutrition and cand.kcal is None):continue
        if not candidate_allowed(cand,result.meal_variant):continue
        trial=dict(result.assignment);trial[target_key]=cand.id
        if hard_rule_violations(result.dates,result.rules,trial,result.candidates):continue
        p,_,_,_=evaluate(result.dates,result.structure,result.rules,trial,list(result.candidates.values()),result.kcal_min,result.kcal_max);scored.append((cand,p,cand.id==current))
    scored.sort(key=lambda x:(x[1],x[0].name));return scored if limit is None else scored[:limit]
