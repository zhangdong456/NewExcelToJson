using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using Newtonsoft.Json;
using UnityEngine;


/*
 *重点！！！！！！！！！！！！！！！！！！！！！！！
 *
 *该代码属于导表工具自动生成
 *切勿修改该代码 否则下次导表会被覆盖
 */


[System.Serializable]
[Config]
public partial class RoleCategory:IConfigCategory
{
	[JsonIgnore]
	public static RoleCategory Inst;


	private Dictionary<int, Role> dict = new Dictionary<int, Role>();
	[SerializeField]
	public List<Role> dataList= new List<Role>();

	public RoleCategory()
	{
		Inst = this;
	}


	public override void EndInit()
	{
		foreach (Role config in dataList)
		{
			config.EndInit();
			this.dict.Add(config.Id,config);
		}
		this.AfterEndInit(); 
	}
	public Role GetOne(int id)
	{
		if (dict.TryGetValue(id,out Role c))
		{
			return c;
		}

		return null;
	}
	public bool isContain(int id)
	{
		return dict.ContainsKey(id);
	}
	public Dictionary<int, Role> GetDict()
	{
		return dict;
	}
}



[System.Serializable]
public partial class Role:IConfig
{
	public int Id; // id
	public string Name; // 名字
	public string Name_short; // 短名字
	public string NickName; // 外号
	public string[] Cast; // 特性
	public int Ability_1; // 体质
	public int Ability_3; // 消耗
	public int Ability_7; // 信用
	public int Ability_2; // 智力
	public int Ability_4; // 观察力
	public int Ability_5; // 口才
	public int Ability_6; // 会计
	public int Ability_8; // 抗压
	public int Ability_9; // 机械
	public int Ability_10; // 生产
	public int Ability_11; // 商业
	public int Ability_12; // 军政
	public int Initial; // 是否初始角色
	public int Pay; // 招募花费
	public int Random; // 是否随机特性
	public int Probability; // 权重
	public int Refresh; // 刷新限制
	public int Gender; // 性别
	public string CharIcon; // 头像
	public int SepName; // 国籍ID
	public string CharImage; // 形象
}

