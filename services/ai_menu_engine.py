"""團膳 AI 菜單規則引擎：排菜、評分、換菜與結果重算。"""
from __future__ import annotations
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
import random
from sqlalchemy.orm import joinedload
from models import KitchenRecipe, KitchenRecipeIngredient
from services.nutrition import recipe_nutrition
from services.food_servings import GROUP_ORDER, finalize_day_servings, merge_recipe_servings, recipe_servings, serving_deviation, targets_for_kcal

CATEGORY_ORDER=("主食","主菜","副菜","青菜","湯品","點心")
DEFAULT_STRUCTURE={"主食":1,"主菜":1,"副菜":2,"青菜":1,"湯品":1,"點心":0}
STRUCTURE_LIMITS={"主食":2,"主菜":2,"副菜":4,"青菜":2,"湯品":2,"點心":2}

@dataclass
class Rules:
    recipe_repeat_days:int=5
    main_repeat_days:int=5
    fish_per_week_min:int=1
    fried_per_week_max:int=1
    sweet_soup_per_week_max:int=1
    prefer_kcal_in_range:bool=True
    exclude_incomplete_nutrition:bool=True
    @classmethod
    def from_dict(cls,data):
        data=data or {}; r=cls()
        legacy=data.get("repeat_days")
        if legacy is not None:r.recipe_repeat_days=r.main_repeat_days=max(0,int(legacy))
        for k in ("recipe_repeat_days","main_repeat_days","fish_per_week_min","fried_per_week_max","sweet_soup_per_week_max"):
            if data.get(k) is not None:
                try:setattr(r,k,max(0,int(data[k])))
                except (TypeError,ValueError):pass
        for k in ("prefer_kcal_in_range","exclude_incomplete_nutrition"):
            if k in data:setattr(r,k,bool(data[k]))
        return r
    def to_dict(self):return dict(self.__dict__)

@dataclass(frozen=True)
class Candidate:
    id:int; name:str; category:str; kcal:int|None; tags:frozenset[str]=frozenset(); servings:tuple=()
    def serving_dict(self):return dict(self.servings)

@dataclass
class DayResult:
    service_date:date; kcal:int|None; warnings:list[str]=field(default_factory=list); ok_messages:list[str]=field(default_factory=list); servings:dict[str,float]=field(default_factory=dict)

@dataclass
class PlanResult:
    dates:list[date]; structure:dict; rules:Rules; assignment:dict; candidates:dict[int,Candidate]; penalty:float; breakdown:dict[str,float]; days:dict[date,DayResult]; warnings:list[str]; kcal_min:int|None=None; kcal_max:int|None=None
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
    lines=[f"同一道菜 {r.recipe_repeat_days} 天內不重複",f"主菜 {r.main_repeat_days} 天內不重複",f"每週至少 {r.fish_per_week_min} 次魚類",f"每週最多 {r.fried_per_week_max} 次炸物",f"每週最多 {r.sweet_soup_per_week_max} 次甜湯"]
    if r.prefer_kcal_in_range:lines.append("每日總熱量與六大類份數盡量接近供餐對象基準")
    if r.exclude_incomplete_nutrition:lines.append("排除營養資料不完整的菜色")
    return lines

def load_candidates():
    rows=(KitchenRecipe.query.options(joinedload(KitchenRecipe.ingredients).joinedload(KitchenRecipeIngredient.ingredient),joinedload(KitchenRecipe.tags)).filter(KitchenRecipe.active.is_(True)).order_by(KitchenRecipe.name).all())
    out=[]
    for r in rows:
        cat=(r.category or "").strip()
        if cat not in CATEGORY_ORDER:continue
        n=recipe_nutrition(r);tags=frozenset(t.tag for t in (r.tags or []) if getattr(t,"source","manual")=="manual");sv=recipe_servings(r).rounded()
        out.append(Candidate(r.id,r.name,cat,n.kcal_int,tags,tuple((k,sv.get(k,0.0)) for k in GROUP_ORDER)))
    return out

def evaluate(dates,structure,rules,assignment,candidates,kcal_min=None,kcal_max=None):
    by_id={c.id:c for c in candidates};breakdown=defaultdict(float);day_results={};warnings=[];uses=defaultdict(list);main_uses=defaultdict(list);week_tags=defaultdict(Counter)
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
        if cnt["fish"]<rules.fish_per_week_min:breakdown["fish"]+=(rules.fish_per_week_min-cnt["fish"])*2500;warnings.append(f"⚠ {week[0]}年第{week[1]}週：魚類 {cnt['fish']} 次，未達 {rules.fish_per_week_min} 次")
        if cnt["fried"]>rules.fried_per_week_max:breakdown["fried"]+=(cnt["fried"]-rules.fried_per_week_max)*1500
        if cnt["sweet_soup"]>rules.sweet_soup_per_week_max:breakdown["sweet_soup"]+=(cnt["sweet_soup"]-rules.sweet_soup_per_week_max)*800
    return sum(breakdown.values()),dict(breakdown),day_results,warnings

def generate(start,end,structure,rules,kcal_min=None,kcal_max=None,include_weekends=False,seed=20260913,locked=None):
    rng=random.Random(seed);dates=service_dates(start,end,include_weekends);candidates=load_candidates();by_id={c.id:c for c in candidates};pools=defaultdict(list)
    for c in candidates:
        if rules.exclude_incomplete_nutrition and c.kcal is None:continue
        pools[c.category].append(c)
    for p in pools.values():rng.shuffle(p)
    assignment=dict(locked or {});last={};usage=Counter()
    for i,d in enumerate(dates):
        day_seen={assignment[k] for k in assignment if k[0]==i and assignment[k]}
        for c in CATEGORY_ORDER:
            for s in range(1,structure.get(c,0)+1):
                key=(i,c,s)
                if key in assignment and assignment[key]:continue
                best=None;bestscore=None
                for cand in pools.get(c,[]):
                    if cand.id in day_seen:continue
                    gap=i-last.get(cand.id,-999);score=usage[cand.id]*60+rng.random()
                    if gap<rules.recipe_repeat_days:score+=6000
                    if c=="主菜" and gap<rules.main_repeat_days:score+=4000
                    if bestscore is None or score<bestscore:best,bestscore=cand,score
                assignment[key]=best.id if best else None
                if best:day_seen.add(best.id);last[best.id]=i;usage[best.id]+=1
    locked_keys=set((locked or {}).keys());current_penalty,_,_,_=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max);movable=[key for key in assignment if key not in locked_keys]
    for _pass in range(2):
        improved=False;rng.shuffle(movable)
        for key in movable:
            i,cat,_slot=key;current=assignment.get(key);same={rid for k,rid in assignment.items() if k[0]==i and k!=key and rid};trial_pool=[cand for cand in pools.get(cat,[]) if cand.id not in same];rng.shuffle(trial_pool);best_id=current;best_penalty=current_penalty
            for cand in trial_pool[:16]:
                if cand.id==current:continue
                trial=dict(assignment);trial[key]=cand.id;p,_,_,_=evaluate(dates,structure,rules,trial,candidates,kcal_min,kcal_max)
                if p+0.01<best_penalty:best_id,best_penalty=cand.id,p
            if best_id!=current:assignment[key]=best_id;current_penalty=best_penalty;improved=True
        if not improved:break
    penalty,breakdown,days_report,warnings=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max)
    return PlanResult(dates,structure,rules,assignment,by_id,penalty,breakdown,days_report,warnings,kcal_min,kcal_max)

def rebuild(dates,structure,rules,assignment,kcal_min=None,kcal_max=None):
    candidates=load_candidates();penalty,breakdown,days_report,warnings=evaluate(dates,structure,rules,assignment,candidates,kcal_min,kcal_max);return PlanResult(dates,structure,rules,assignment,{c.id:c for c in candidates},penalty,breakdown,days_report,warnings,kcal_min,kcal_max)

def replacement_options(result,target_key,limit=40):
    i,cat,slot=target_key;same={rid for key,rid in result.assignment.items() if key[0]==i and key!=target_key and rid};current=result.assignment.get(target_key);scored=[]
    for cand in result.candidates.values():
        if cand.category!=cat or cand.id in same or (result.rules.exclude_incomplete_nutrition and cand.kcal is None):continue
        trial=dict(result.assignment);trial[target_key]=cand.id;p,_,_,_=evaluate(result.dates,result.structure,result.rules,trial,list(result.candidates.values()),result.kcal_min,result.kcal_max);scored.append((cand,p,cand.id==current))
    scored.sort(key=lambda x:(x[1],x[0].name));return scored[:limit]
