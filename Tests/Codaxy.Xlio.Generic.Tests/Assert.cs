using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PetaTest
{
    public partial class Assert
    {
        public static void AreEqualAndHaveSameHash(object o1, object o2)
        {
            if (o1.GetHashCode() != o2.GetHashCode())
                throw new AssertionException("Given objects do not have the same hash code.");
            if (!o1.Equals(o2) || !o2.Equals(o1))
                throw new AssertionException("Given objects are not the same.");
        }
    }
}
