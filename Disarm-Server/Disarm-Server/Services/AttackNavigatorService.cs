using Disarm_Server.Models;
using System.Text.Json;

namespace Disarm_Server.Services;

public class AttackNavigatorService : IAttackNavigatorService
{
    private readonly DisarmTechniqueNameWrapper _disarmTechniqueNameWrapper;

    public AttackNavigatorService(DisarmTechniqueNameWrapper disarmTechniqueNameWrapper)
    {
        _disarmTechniqueNameWrapper = disarmTechniqueNameWrapper;
    }
    public AttackNavigatorLayer GetNavigatorLayerBasedOnTechniqueIds(string[] techniqueIds)
    {
        var disarmData = _disarmTechniqueNameWrapper.Techniques
            .ToDictionary(t => t.Id, t => t.PhaseName);

        var ids = techniqueIds.ToLookup(id => id.Contains('.'));

        var subIds = ids[true].ToList();
        var parentIds = ids[false].ToList();

        var idDictionary = subIds
            .GroupBy(id => id.Split('.').First())
            .ToDictionary(t => t.Key, t => t.ToList());

        foreach (var id in parentIds)
        {
            idDictionary.TryAdd(id, []);
        }

        var attackNavigatorLayer = new AttackNavigatorLayer
        {
            Techniques = []
        };

        foreach (var id in idDictionary)
        {
            var technique = new Technique
            {
                TechniqueID = id.Key,
                Tactic = disarmData[id.Key],
                Score = 1,
                Color = "#e60d0d",
                Comment = "",
                Enabled = true,
                Metadata = [],
                Links = [],
                ShowSubtechniques = id.Value.Count > 0,
            };

            foreach (var subId in id.Value)
            {
                var subTechnique = new Technique
                {
                    TechniqueID = subId,
                    Tactic = disarmData[subId],
                    Score = 1,
                    Color = "#e60d0d",
                    Comment = "",
                    Enabled = true,
                    Metadata = [],
                    Links = [],
                    ShowSubtechniques = false,
                };
                attackNavigatorLayer.Techniques.Add(subTechnique);
            }
            attackNavigatorLayer.Techniques.Add(technique);
        }


        return attackNavigatorLayer;
    }


}
