// Guids.cs
// MUST match guids.h
using System;

namespace NielsV.NielsV_AttachToAppPool
{
    static class GuidList
    {
        public const string guidNielsV_AttachToAppPoolPkgString = "ce08f90c-07be-4d3a-bb11-c19b1e01dcf8";
        public const string guidNielsV_AttachToAppPoolCmdSetString = "4c0ff035-ad9b-4063-aa7d-a2989fee6a4b";

        public static readonly Guid guidNielsV_AttachToAppPoolCmdSet = new Guid(guidNielsV_AttachToAppPoolCmdSetString);
    };
}