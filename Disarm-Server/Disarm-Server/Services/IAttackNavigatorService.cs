using Disarm_Server.Models;

namespace Disarm_Server.Services;

public interface IAttackNavigatorService
{
    AttackNavigatorLayer GetNavigatorLayerBasedOnTechniqueIds(string[] techniqueIds);
}