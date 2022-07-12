Students were "given" an Office 365 license to use with their student account when they enrolled, but due to the fact only a limited number of administrators could do this, and there were thousands of students enrolling every year, it was relegated to a request, i.e. the student had to request the license first before it would be assigned since maintaining the licenses manually proved impossible.

This script resolved the issue by retrieving a list of active students and then adjusting license assignments accordingly.

Uses an application registration in Azure AD along with a certificate for authentication to the Graph API.

Not the fastest thing in the world, but still a lot faster than doing them by hand.

Won't really work as-is due to stripping identifying information out of it since it was built just for that environment.
